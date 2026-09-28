import * as fs from "fs";
import * as path from "path";
import * as vscode from "vscode";
import {
    LanguageClient,
    LanguageClientOptions,
    ServerOptions,
    TransportKind
} from "vscode-languageclient/node";
import { splitCommandLine } from "./commandLine";
import { ScriptEncodings } from "./encoding";
import { CodeLibrary } from "./library";
import { PlatformStatus, ScriptRunner, activeScript, resolveScriptHost, scriptText, togglePlatform } from "./run";

/**
 * Thin client: syntax highlighting comes from the TextMate grammar (works on
 * every OS); IntelliSense comes from VbsTools.LanguageServer.exe, which reads
 * COM type libraries from the Windows registry and therefore runs on Windows only.
 */

let client: LanguageClient | undefined;
let output: vscode.OutputChannel | undefined;
let runner: ScriptRunner | undefined;

export async function activate(context: vscode.ExtensionContext): Promise<void> {
    output = vscode.window.createOutputChannel("VBScript Language Server");
    context.subscriptions.push(output);

    runner = new ScriptRunner(context);
    const scripts = runner;

    context.subscriptions.push(
        scripts,
        new PlatformStatus(),
        vscode.commands.registerCommand("vbscript.run", () => scripts.run(false)),
        vscode.commands.registerCommand("vbscript.runWithArgs", () => scripts.run(true)),
        vscode.commands.registerCommand("vbscript.stop", () => scripts.stop()),
        vscode.commands.registerCommand("vbscript.togglePlatform", togglePlatform),
        vscode.commands.registerCommand("vbscript.debugWithArgs", () => debugWithArguments(scripts)),
        // Code lens "N references": the server sends plain JSON; showReferences wants VS Code types.
        vscode.commands.registerCommand("vbscript.showReferences",
            (uri: string, position: { line: number; character: number }, locations: { uri: string; range: LspRange }[]) =>
                vscode.commands.executeCommand("editor.action.showReferences",
                    vscode.Uri.parse(uri),
                    new vscode.Position(position.line, position.character),
                    locations.map(l => new vscode.Location(vscode.Uri.parse(l.uri), toRange(l.range))))),
        vscode.commands.registerCommand("vbscript.restartServer", async () => {
            await stopClient();
            await startClient(context);
        }),
        vscode.debug.registerDebugConfigurationProvider("vbscript", new VbsDebugConfigurationProvider(context, scripts)),
        vscode.debug.registerDebugAdapterDescriptorFactory("vbscript", new VbsDebugAdapterFactory(context)),
        vscode.debug.registerDebugAdapterTrackerFactory("vbscript", new VbsDebugTraceFactory()),
        ...new CodeLibrary(context).register(),
        ...new ScriptEncodings().register(),
        vscode.commands.registerCommand("vbscript.openDocumentation", (page?: string) => openDocumentation(context, page))
    );

    await startClient(context);
}

/** The user guide (docs/<language>/*.md) in the Markdown preview: Simplified Chinese or English. */
async function openDocumentation(context: vscode.ExtensionContext, page = "index.md"): Promise<void> {
    const display = vscode.env.language.toLowerCase();
    const language = display.startsWith("zh") && !/tw|hk|mo|hant/.test(display) ? "zh-cn" : "en";
    const uri = vscode.Uri.joinPath(context.extensionUri, "docs", language, page);

    await vscode.commands.executeCommand("markdown.showPreview", uri);
}


interface LspRange {
    start: { line: number; character: number };
    end: { line: number; character: number };
}

function toRange(range: LspRange): vscode.Range {
    return new vscode.Range(range.start.line, range.start.character, range.end.line, range.end.character);
}

export async function deactivate(): Promise<void> {
    await stopClient();
}


function resolveServerPath(context: vscode.ExtensionContext): string | undefined {
    const configured = vscode.workspace.getConfiguration("vbscript").get<string>("languageServer.path", "").trim();

    const candidates = configured
        ? [configured]
        : [context.asAbsolutePath(path.join("server", "VbsTools.LanguageServer.exe"))];

    return candidates.find(p => fs.existsSync(p));
}


async function startClient(context: vscode.ExtensionContext): Promise<void> {
    if (process.platform !== "win32") {
        output?.appendLine(vscode.l10n.t("IntelliSense needs Windows (COM type libraries are read from the registry). Syntax highlighting still works."));
        return;
    }

    const serverPath = resolveServerPath(context);

    if (!serverPath) {
        const message = vscode.l10n.t("VBScript language server not found. Build it with build-server.ps1 or set 'vbscript.languageServer.path'.");
        output?.appendLine(message);
        void vscode.window.showWarningMessage(message);
        return;
    }

    const trace = vscode.workspace.getConfiguration("vbscript").get<boolean>("languageServer.trace", false);

    const serverOptions: ServerOptions = {
        command: serverPath,
        args: trace ? ["--trace"] : [],
        transport: TransportKind.stdio
    };

    const clientOptions: LanguageClientOptions = {
        documentSelector: [
            { scheme: "file", language: "vbscript" },
            { scheme: "untitled", language: "vbscript" }
        ],
        outputChannel: output,
        // Sends the "vbscript" settings to the server (workspace/didChangeConfiguration).
        synchronize: {
            configurationSection: "vbscript",
            // Type definition files: the server reloads them when they change.
            fileEvents: vscode.workspace.createFileSystemWatcher("**/*.vbstypes.json")
        }
    };

    client = new LanguageClient("vbscript", "VBScript Language Server", serverOptions, clientOptions);

    try {
        await client.start();
        output?.appendLine(`Started ${serverPath}`);
    } catch (error) {
        output?.appendLine(`Failed to start ${serverPath}: ${error}`);
        void vscode.window.showErrorMessage(vscode.l10n.t("VBScript language server failed to start: {0}", String(error)));
        client = undefined;
    }
}


async function stopClient(): Promise<void> {
    if (client) {
        const running = client;
        client = undefined;
        await running.stop();
    }
}


/** Asks for the arguments, then debugs the active script with them. */
async function debugWithArguments(scripts: ScriptRunner): Promise<void> {
    const document = await activeScript();

    if (!document) {
        return;
    }

    const commandLine = await scripts.promptArguments(document.fileName);

    if (commandLine === undefined) {
        return;
    }

    await vscode.debug.startDebugging(vscode.workspace.getWorkspaceFolder(document.uri), {
        type: "vbscript",
        request: "launch",
        name: "Debug VBScript",
        program: document.fileName,
        args: splitCommandLine(commandLine)
    });
}


/** Fills in a launch configuration: F5 without launch.json debugs the active script. */
class VbsDebugConfigurationProvider implements vscode.DebugConfigurationProvider {
    constructor(private readonly context: vscode.ExtensionContext, private readonly scripts: ScriptRunner) { }

    resolveDebugConfiguration(
        _folder: vscode.WorkspaceFolder | undefined,
        config: vscode.DebugConfiguration): vscode.ProviderResult<vscode.DebugConfiguration> {

        if (process.platform !== "win32") {
            void vscode.window.showErrorMessage(vscode.l10n.t("Debugging VBScript requires Windows."));
            return undefined;
        }

        if (!config.type && !config.request && !config.name) {
            const editor = vscode.window.activeTextEditor;

            if (!editor || editor.document.languageId !== "vbscript") {
                return undefined;   // let VS Code offer launch.json
            }

            config.type = "vbscript";
            config.name = "Debug VBScript";
            config.request = "launch";
            config.program = "${file}";
        }

        config.program ??= "${file}";

        const platform = (config.platform ?? vscode.workspace.getConfiguration("vbscript").get<string>("run.platform", "x64")) === "x86" ? "x86" : "x64";
        config.platform = platform;

        // Without debugging the script runs with vbscript.run.host (see below).
        if (!config.noDebug) {
            const host = resolveScriptHost(this.context, platform);

            if (!host) {
                return undefined;
            }

            config.hostExecutable = host;
        }

        return config;
    }

    async resolveDebugConfigurationWithSubstitutedVariables(
        _folder: vscode.WorkspaceFolder | undefined,
        config: vscode.DebugConfiguration): Promise<vscode.DebugConfiguration | undefined> {

        // Debug what is on screen: save the script first.
        const document = vscode.workspace.textDocuments.find(d => d.fileName.toLowerCase() === String(config.program).toLowerCase());

        if (document?.isDirty && !(await document.save())) {
            return undefined;
        }

        // No "args" in the configuration: the ones remembered by "Run with Arguments…".
        if (config.args === undefined && typeof config.program === "string") {
            const remembered = this.scripts.getArguments(config.program);

            if (remembered.trim()) {
                config.args = splitCommandLine(remembered);
            }
        }

        // No "console" in the configuration: like Run Script (vbscript.run.console) — a script
        // that reads WScript.StdIn gets a terminal, the Debug Console has no input.
        if (config.console === undefined && typeof config.program === "string") {
            const mode = vscode.workspace.getConfiguration("vbscript").get<string>("run.console", "auto");

            if (mode === "terminal" || (mode === "auto" && /\bStdIn\b/i.test(scriptText(config.program)))) {
                config.console = "integratedTerminal";
            }
        }

        // "Run Without Debugging" (Ctrl+F5 outside a VBScript editor, the Run button's menu):
        // run like "VBScript: Run Script" — with vbscript.run.host, output in the VBScript
        // panel, errors in Problems — instead of starting the debugger.
        if (config.noDebug && typeof config.program === "string") {
            this.scripts.runFile(config.program, Array.isArray(config.args) ? config.args.map(String) : [],
                { platform: config.platform, cwd: typeof config.cwd === "string" ? config.cwd : undefined });
            return undefined;   // no debug session
        }

        return config;
    }
}


/** "vbscript.debug.trace": every message between VS Code and the debug adapter, in an output channel. */
class VbsDebugTraceFactory implements vscode.DebugAdapterTrackerFactory {
    private channel: vscode.OutputChannel | undefined;

    createDebugAdapterTracker(): vscode.ProviderResult<vscode.DebugAdapterTracker> {
        if (!vscode.workspace.getConfiguration("vbscript").get<boolean>("debug.trace", false)) {
            return undefined;
        }

        const channel = this.channel ??= vscode.window.createOutputChannel("VBScript Debug Trace");
        const time = () => new Date().toISOString().substring(11, 23);

        return {
            onWillReceiveMessage: message => channel.appendLine(`${time()} → ${JSON.stringify(message)}`),
            onDidSendMessage: message => channel.appendLine(`${time()} ← ${JSON.stringify(message)}`),
            onError: error => channel.appendLine(`${time()} ! ${error.message}`),
            onExit: (code, signal) => channel.appendLine(`${time()} adapter exited (${code ?? signal})`)
        };
    }
}


/** The debug adapter is the language server executable in "--dap" mode. */
class VbsDebugAdapterFactory implements vscode.DebugAdapterDescriptorFactory {
    constructor(private readonly context: vscode.ExtensionContext) { }

    createDebugAdapterDescriptor(): vscode.ProviderResult<vscode.DebugAdapterDescriptor> {
        const server = resolveServerPath(this.context);

        if (!server) {
            throw new Error(vscode.l10n.t("VbsTools.LanguageServer.exe not found (it is also the debug adapter). Build it with build-server.ps1 or set 'vbscript.languageServer.path'."));
        }

        return new vscode.DebugAdapterExecutable(server, ["--dap"]);
    }
}
