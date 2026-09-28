import { ChildProcess, execSync, spawn } from "child_process";
import * as fs from "fs";
import * as path from "path";
import * as vscode from "vscode";
import { quoteArgument, splitCommandLine } from "./commandLine";
import { scriptEditor } from "./walkthrough";

/**
 * "VBScript: Run Script" (Ctrl+F5) and "Run with Arguments…" (Ctrl+Shift+F5).
 *
 * Output goes to the "VBScript" output panel; a runtime or syntax error is shown in
 * the Problems panel and clicking it opens the line. Scripts that read WScript.StdIn
 * run in a terminal instead (the output panel has no input); see vbscript.run.console.
 * Arguments are remembered per script and also used by F5 (debugging without launch.json).
 */
export class ScriptRunner implements vscode.Disposable {
    private readonly output = vscode.window.createOutputChannel("VBScript");
    private readonly problems = vscode.languages.createDiagnosticCollection("vbscript-run");
    private readonly disposables: vscode.Disposable[] = [];
    private terminal: vscode.Terminal | undefined;
    private child: ChildProcess | undefined;

    constructor(private readonly context: vscode.ExtensionContext) {
        this.disposables.push(
            this.output,
            this.problems,
            vscode.window.onDidCloseTerminal(t => {
                if (t === this.terminal) {
                    this.terminal = undefined;
                }
            }),
            // An edited script: its last run's errors may no longer apply.
            vscode.workspace.onDidChangeTextDocument(e => {
                if (e.contentChanges.length > 0 && this.problems.has(e.document.uri)) {
                    this.problems.delete(e.document.uri);
                }
            })
        );
    }

    dispose(): void {
        this.stop();
        this.disposables.forEach(d => d.dispose());
    }


    // ============================================================
    // Arguments (remembered per script)
    // ============================================================

    getArguments(file: string): string {
        const all = this.context.workspaceState.get<Record<string, string>>("vbscript.arguments", {});
        return all[file.toLowerCase()] ?? "";
    }

    private async setArguments(file: string, commandLine: string): Promise<void> {
        const all = { ...this.context.workspaceState.get<Record<string, string>>("vbscript.arguments", {}) };

        if (commandLine.trim()) {
            all[file.toLowerCase()] = commandLine;
        } else {
            delete all[file.toLowerCase()];
        }

        await this.context.workspaceState.update("vbscript.arguments", all);
    }

    /** Asks for the arguments (last ones pre-filled); undefined = cancelled. */
    async promptArguments(file: string): Promise<string | undefined> {
        const value = await vscode.window.showInputBox({
            title: vscode.l10n.t("Arguments for {0}", path.basename(file)),
            prompt: vscode.l10n.t("Command line passed to the script (WScript.Arguments). Quote arguments that contain spaces. Remembered for this script, also for F5."),
            value: this.getArguments(file),
            placeHolder: 'in.txt /out:"result file.txt" /verbose'
        });

        if (value !== undefined) {
            await this.setArguments(file, value);
        }

        return value;
    }


    // ============================================================
    // Run
    // ============================================================

    get isRunning(): boolean {
        return this.child !== undefined;
    }

    async run(withPrompt: boolean): Promise<void> {
        const document = await activeScript();

        if (!document) {
            return;
        }

        const file = document.fileName;
        let commandLine = this.getArguments(file);

        if (withPrompt) {
            const entered = await this.promptArguments(file);

            if (entered === undefined) {
                return;
            }

            commandLine = entered;
        }

        this.runFile(file, splitCommandLine(commandLine));
    }


    /**
     * Runs a script with the configured host (vbscript.run.host). Also used for
     * "Run Without Debugging" (a launch with noDebug), so both run the same way.
     */
    runFile(file: string, args: string[], options: { platform?: string; cwd?: string } = {}): void {
        const config = vscode.workspace.getConfiguration("vbscript");
        const platform = (options.platform ?? config.get<string>("run.platform", "x64")) === "x86" ? "x86" : "x64";
        const host = config.get<string>("run.host", "vbstools");
        const vbstools = host === "vbstools";
        const executable = vbstools ? resolveScriptHost(this.context, platform) : resolveCScript(platform);

        if (!executable) {
            return;
        }

        const cwd = options.cwd || path.dirname(file);
        const console = config.get<string>("run.console", "auto");
        const useTerminal = console === "terminal" || (console === "auto" && /\bStdIn\b/i.test(scriptText(file)));

        if (useTerminal) {
            this.runInTerminal(executable, file, args, cwd);
        } else {
            this.runInOutput(executable, file, args, vbstools, platform, host, cwd);
        }
    }


    private runInTerminal(executable: string, file: string, args: string[], cwd: string): void {
        if (!this.terminal) {
            // cmd.exe shows the OEM code page output of both hosts correctly (CJK included).
            this.terminal = vscode.window.createTerminal({ name: "VBScript", shellPath: "cmd.exe" });
        }

        const commandLine = [executable, "//NoLogo", file, ...args].map(quoteArgument).join(" ");

        this.terminal.show(true);
        this.terminal.sendText(`cd /d ${quoteArgument(cwd)} && ${commandLine}`);
    }


    private runInOutput(executable: string, file: string, args: string[], vbstools: boolean, platform: string, host: string, cwd: string): void {
        this.stop();
        this.problems.clear();

        const output = this.output;
        output.clear();
        output.show(true);

        const shownArgs = args.length > 0 ? " " + args.map(quoteArgument).join(" ") : "";
        output.appendLine(`[Running] ${path.basename(file)}${shownArgs}   (${host === "vbstools" ? "VbsTools Script Host" : "cscript"}, ${platform})`);

        const started = Date.now();
        // VbsTools host: UTF-8, and stderr into stdout — two pipes would lose the order of the two.
        const hostArgs = ["//NoLogo", ...(vbstools ? ["//UTF8", "//ErrToOut"] : []), file, ...args];

        const child = spawn(executable, hostArgs, { cwd, windowsHide: true });
        this.child = child;
        void vscode.commands.executeCommand("setContext", "vbscript.running", true);

        // The script has no input here (WScript.StdIn is at its end); scripts that read it run in the terminal.
        child.stdin?.end();

        const decoder = new TextDecoder(vbstools ? "utf-8" : oemEncoding());
        let captured = "";

        const onData = (data: Buffer) => {
            const text = decoder.decode(data, { stream: true });
            captured += text;
            output.append(text);
        };

        child.stdout?.on("data", onData);
        child.stderr?.on("data", onData);

        child.on("error", error => {
            output.appendLine(`[Error] ${error.message}`);
            this.finished(child);
        });

        child.on("close", (code, signal) => {
            const tail = decoder.decode();

            if (tail) {
                captured += tail;
                output.append(tail);
            }

            const seconds = ((Date.now() - started) / 1000).toFixed(1);
            const how = signal ? "stopped" : `exit code ${code}`;
            output.appendLine(`${captured.endsWith("\n") || captured.length === 0 ? "" : "\n"}[Done] ${how}, ${seconds} s`);

            this.reportErrors(captured, file);
            this.finished(child);
        });
    }


    private finished(child: ChildProcess): void {
        if (this.child === child) {
            this.child = undefined;
            void vscode.commands.executeCommand("setContext", "vbscript.running", false);
        }
    }


    stop(): void {
        if (this.child) {
            const child = this.child;

            try {
                child.kill();
            } catch {
                // already ended
            }

            this.finished(child);
        }
    }


    /** "C:\x\a.vbs(3, 5) Microsoft VBScript runtime error: …" → Problems panel. */
    private reportErrors(output: string, file: string): void {
        const byFile = new Map<string, vscode.Diagnostic[]>();
        const pattern = /^(.+?)\((\d+),\s*(\d+)\)\s+(.+)$/gm;
        let match: RegExpExecArray | null;

        while ((match = pattern.exec(output)) !== null) {
            const errorFile = match[1].trim();

            if (!fs.existsSync(errorFile)) {
                continue;
            }

            const line = Math.max(0, parseInt(match[2], 10) - 1);
            const column = Math.max(0, parseInt(match[3], 10) - 1);
            const message = match[4].trim();

            const diagnostic = new vscode.Diagnostic(new vscode.Range(line, column, line, column + 1), message, vscode.DiagnosticSeverity.Error);
            diagnostic.source = "run";

            const key = path.resolve(errorFile).toLowerCase();
            byFile.set(key, [...(byFile.get(key) ?? []), diagnostic]);
        }

        for (const [key, diagnostics] of byFile) {
            const target = key === file.toLowerCase() ? vscode.Uri.file(file) : vscode.Uri.file(key);
            this.problems.set(target, diagnostics);
        }

        // Put the caret on the error, like VbsTools Studio.
        const first = byFile.get(file.toLowerCase())?.[0];

        if (first) {
            const editor = vscode.window.visibleTextEditors.find(e => e.document.fileName.toLowerCase() === file.toLowerCase());

            if (editor) {
                editor.selection = new vscode.Selection(first.range.start, first.range.start);
                editor.revealRange(first.range, vscode.TextEditorRevealType.InCenterIfOutsideViewport);
            }
        }
    }
}


// ============================================================
// Status bar: platform x64 / x86
// ============================================================

export class PlatformStatus implements vscode.Disposable {
    private readonly item = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Right, 100);
    private readonly disposables: vscode.Disposable[] = [];

    constructor() {
        this.item.command = "vbscript.togglePlatform";
        this.disposables.push(
            this.item,
            vscode.window.onDidChangeActiveTextEditor(() => this.update()),
            vscode.workspace.onDidChangeConfiguration(e => {
                if (e.affectsConfiguration("vbscript.run")) {
                    this.update();
                }
            })
        );
        this.update();
    }

    private update(): void {
        const editor = vscode.window.activeTextEditor;

        if (!editor || editor.document.languageId !== "vbscript") {
            this.item.hide();
            return;
        }

        const config = vscode.workspace.getConfiguration("vbscript");
        const platform = config.get<string>("run.platform", "x64");
        const host = config.get<string>("run.host", "vbstools") === "vbstools" ? "VbsTools" : "cscript";

        this.item.text = `$(chip) VBScript ${platform}`;
        this.item.tooltip = platform === "x86"
            ? vscode.l10n.t("Scripts run and debug as 32-bit ({0}). Click to switch.", host)
            : vscode.l10n.t("Scripts run and debug as 64-bit ({0}). Click to switch.", host);
        this.item.show();
    }

    dispose(): void {
        this.disposables.forEach(d => d.dispose());
    }
}


/** Switches vbscript.run.platform where it is set (workspace if set there, else user settings). */
export async function togglePlatform(): Promise<void> {
    const config = vscode.workspace.getConfiguration("vbscript");
    const next = config.get<string>("run.platform", "x64") === "x86" ? "x64" : "x86";
    const inspected = config.inspect<string>("run.platform");

    const target = inspected?.workspaceFolderValue !== undefined ? vscode.ConfigurationTarget.WorkspaceFolder
        : inspected?.workspaceValue !== undefined ? vscode.ConfigurationTarget.Workspace
            : vscode.ConfigurationTarget.Global;

    await config.update("run.platform", next, target);
    void vscode.window.setStatusBarMessage(`VBScript: ${next === "x86" ? "32-bit (x86)" : "64-bit (x64)"}`, 3000);
}


// ============================================================
// Helpers
// ============================================================

/** The script's text as open in the editor, else from disk ("" when unreadable). */
export function scriptText(file: string): string {
    const open = vscode.workspace.textDocuments.find(d => d.fileName.toLowerCase() === file.toLowerCase());

    if (open) {
        return open.getText();
    }

    try {
        return fs.readFileSync(file, "latin1");
    } catch {
        return "";
    }
}

/** The active VBScript document, saved (untitled scripts must be saved first). */
export async function activeScript(): Promise<vscode.TextDocument | undefined> {
    const editor = scriptEditor();

    if (!editor || editor.document.languageId !== "vbscript") {
        return undefined;
    }

    if (process.platform !== "win32") {
        void vscode.window.showErrorMessage(vscode.l10n.t("Running VBScript requires Windows."));
        return undefined;
    }

    const document = editor.document;

    if (document.isUntitled) {
        void vscode.window.showInformationMessage(vscode.l10n.t("Save the script first."));
        return undefined;
    }

    if (document.isDirty && !(await document.save())) {
        return undefined;
    }

    return document;
}


export function resolveCScript(platform: string): string | undefined {
    const windows = process.env.SystemRoot ?? "C:\\Windows";

    // VS Code is a 64-bit process: System32 is the real 64-bit directory.
    const cscript = path.join(windows, platform === "x86" ? "SysWOW64" : "System32", "cscript.exe");

    if (!fs.existsSync(cscript)) {
        void vscode.window.showErrorMessage(vscode.l10n.t("{0} not found. VBScript may be disabled on this Windows installation.", cscript));
        return undefined;
    }

    return cscript;
}


export function resolveScriptHost(context: vscode.ExtensionContext, platform: string): string | undefined {
    const configured = vscode.workspace.getConfiguration("vbscript").get<string>("run.scriptHostPath", "").trim();
    const folder = configured || context.asAbsolutePath(path.join("server", "host"));
    const exe = path.join(folder, platform, "VbsTools.ScriptHost.exe");

    if (!fs.existsSync(exe)) {
        void vscode.window.showErrorMessage(
            vscode.l10n.t("{0} not found. Build it with build-server.ps1, set 'vbscript.run.scriptHostPath', or set 'vbscript.run.host' to cscript.", exe));
        return undefined;
    }

    return exe;
}


let oemLabel: string | undefined;

/** TextDecoder label of the console (OEM) code page, used to read cscript's output. */
function oemEncoding(): string {
    if (oemLabel) {
        return oemLabel;
    }

    let codePage = 437;

    try {
        const match = /(\d+)/.exec(execSync("chcp", { encoding: "latin1", windowsHide: true }));

        if (match) {
            codePage = parseInt(match[1], 10);
        }
    } catch {
        // keep the default
    }

    const labels: Record<number, string> = {
        932: "shift_jis", 936: "gbk", 949: "euc-kr", 950: "big5", 65001: "utf-8",
        866: "ibm866", 874: "windows-874", 1250: "windows-1250", 1251: "windows-1251", 1252: "windows-1252"
    };

    const label = labels[codePage] ?? "windows-1252";

    try {
        new TextDecoder(label);
        oemLabel = label;
    } catch {
        oemLabel = "utf-8";
    }

    return oemLabel;
}
