import * as cp from "child_process";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import * as vscode from "vscode";

/**
 * Attach debugging ("request": "attach") needs the Process Debug Manager (pdm.dll) and the Machine Debug Manager
 * (mdm.exe, a Windows service) on the machine, and the user in the local group "Debugger Users". None of them is ours to
 * ship (they belong to Visual Studio / Office), so the extension only checks that they are there, says what is missing
 * and helps to set up what the user already has. See docs/DEBUGGER.md, "Attaching to another application".
 */

const PDM_CLSID = "{78A51822-51F4-11D0-8F20-00805F2CD064}";
const MDM_SERVICE = "MDM";
const GROUP = "Debugger Users";

interface RunResult {
    code: number;
    out: string;
}

interface Check {
    ok: boolean;
    /** Only reported; does not make the environment "not ready". */
    optional?: boolean;
    title: string;
    detail: string;
}

export interface AttachReport {
    checks: Check[];
    mdmRegistered: boolean;
    inGroup: boolean;
    groupExists: boolean;
    addedToGroup: boolean;
    ready: boolean;
}


/**
 * The Windows tools by their full path: another program of the same name earlier on the PATH (Git's whoami, ...) must
 * not answer. A name with a path is used as it is.
 */
function resolveTool(file: string): string {
    if (/[\\/]/.test(file))
        return file;

    const system32 = path.join(process.env.SystemRoot ?? "C:\\Windows", "System32");
    const name = file.toLowerCase().endsWith(".exe") ? file : file + ".exe";

    return name.toLowerCase() === "powershell.exe" ? path.join(system32, "WindowsPowerShell", "v1.0", name) : path.join(system32, name);
}


function run(file: string, args: string[]): Promise<RunResult> {
    return new Promise(resolve => {
        cp.execFile(resolveTool(file), args, { windowsHide: true, encoding: "latin1", timeout: 20000 }, (error, stdout) => {
            const code = error ? (typeof error.code === "number" ? error.code : 1) : 0;
            resolve({ code, out: String(stdout ?? "") });
        });
    });
}


function expandVariables(text: string): string {
    return text.replace(/%([^%]+)%/g, (all, name: string) => process.env[name] ?? all);
}


/** The pdm.dll registered for the 32-bit or the 64-bit view of COM; undefined when there is none or the file is missing. */
async function findPdm(view: "32" | "64"): Promise<string | undefined> {
    const result = await run("reg", ["query", `HKCR\\CLSID\\${PDM_CLSID}\\InprocServer32`, "/ve", `/reg:${view}`]);

    if (result.code !== 0)
        return undefined;

    const match = /REG_(?:EXPAND_)?SZ\s+(.+?)\s*$/m.exec(result.out);
    const file = match ? expandVariables(match[1]) : undefined;

    return file && fs.existsSync(file) ? file : undefined;
}


/** The executable of the Machine Debug Manager service, when it is registered. */
async function findMdm(): Promise<string | undefined> {
    const query = await run("sc", ["query", MDM_SERVICE]);

    if (query.code !== 0)
        return undefined;

    const config = await run("sc", ["qc", MDM_SERVICE]);
    const match = /BINARY_PATH_NAME\s*:\s*(.+?)\s*$/im.exec(config.out);

    return match ? match[1].replace(/^"|"$/g, "") : MDM_SERVICE;
}


function userName(): string {
    return os.userInfo().username;
}


/** Checks everything attach debugging needs from the machine. */
export async function checkAttachEnvironment(): Promise<AttachReport> {
    const [pdm32, pdm64, mdm, token, group] = await Promise.all([
        findPdm("32"),
        findPdm("64"),
        findMdm(),
        run("whoami", ["/groups", "/fo", "csv", "/nh"]),
        run("net", ["localgroup", GROUP])
    ]);

    const inGroup = token.code === 0 && token.out.toLowerCase().includes(GROUP.toLowerCase());
    const groupExists = group.code === 0;
    const addedToGroup = groupExists && group.out.toLowerCase().includes(userName().toLowerCase());
    const checks: Check[] = [];

    checks.push(pdm32
        ? { ok: true, title: vscode.l10n.t("Process Debug Manager (pdm.dll), 32-bit"), detail: pdm32 }
        : {
            ok: false,
            title: vscode.l10n.t("Process Debug Manager (pdm.dll), 32-bit"),
            detail: vscode.l10n.t("Not found. It comes with Visual Studio or Office, not with Windows, and is not part of this extension.")
        });

    checks.push(pdm64
        ? { ok: true, title: vscode.l10n.t("Process Debug Manager (pdm.dll), 64-bit"), detail: pdm64 }
        : {
            ok: false,
            optional: true,
            title: vscode.l10n.t("Process Debug Manager (pdm.dll), 64-bit"),
            detail: vscode.l10n.t("Not found (only needed to debug 64-bit applications).")
        });

    checks.push(mdm
        ? { ok: true, title: vscode.l10n.t("Machine Debug Manager (mdm.exe)"), detail: vscode.l10n.t("Registered as the Windows service {0}: {1}", MDM_SERVICE, mdm) }
        : {
            ok: false,
            title: vscode.l10n.t("Machine Debug Manager (mdm.exe)"),
            detail: vscode.l10n.t("Not registered. Run mdm.exe /regserver as administrator (the Microsoft-signed build).")
        });

    checks.push(inGroup
        ? { ok: true, title: vscode.l10n.t("Group \"{0}\"", GROUP), detail: vscode.l10n.t("{0} is a member.", userName()) }
        : addedToGroup
            ? {
                ok: false,
                title: vscode.l10n.t("Group \"{0}\"", GROUP),
                detail: vscode.l10n.t("{0} was added to the group, but this logon does not have it yet: sign out of Windows and in again.", userName())
            }
            : {
                ok: false,
                title: vscode.l10n.t("Group \"{0}\"", GROUP),
                detail: groupExists
                    ? vscode.l10n.t("{0} is not a member of the group.", userName())
                    : vscode.l10n.t("The group does not exist on this machine.")
            });

    return {
        checks,
        mdmRegistered: mdm !== undefined,
        inGroup,
        groupExists,
        addedToGroup,
        ready: checks.every(c => c.ok || c.optional)
    };
}


// ============================================================
// fixing what the user already has
// ============================================================

/** Runs a program as administrator (the UAC prompt appears); resolves with its exit code. */
function runElevated(file: string, args: string[]): Promise<number> {
    const quote = (text: string) => "'" + text.replace(/'/g, "''") + "'";
    const argumentList = args.length > 0 ? ` -ArgumentList ${args.map(quote).join(",")}` : "";
    const script = `$p = Start-Process -FilePath ${quote(file)}${argumentList} -Verb RunAs -Wait -PassThru; exit $p.ExitCode`;

    return run("powershell.exe", ["-NoProfile", "-NonInteractive", "-Command", script]).then(r => r.code);
}


function runElevatedScript(script: string): Promise<number> {
    const encoded = Buffer.from(script, "utf16le").toString("base64");
    return runElevated("powershell.exe", ["-NoProfile", "-EncodedCommand", encoded]);
}


/** True when the file is signed by Microsoft (an unsigned or foreign mdm.exe is not run as administrator). */
async function isSignedByMicrosoft(file: string): Promise<boolean> {
    const quoted = "'" + file.replace(/'/g, "''") + "'";
    const script = `$s = Get-AuthenticodeSignature -LiteralPath ${quoted}; ` +
        "if ($s.Status -eq 'Valid' -and $s.SignerCertificate.Subject -match 'O=Microsoft Corporation') { exit 0 } else { exit 1 }";

    return (await run("powershell.exe", ["-NoProfile", "-NonInteractive", "-Command", script])).code === 0;
}


async function registerMdm(): Promise<boolean> {
    const picked = await vscode.window.showOpenDialog({
        canSelectMany: false,
        openLabel: vscode.l10n.t("Register this mdm.exe"),
        title: vscode.l10n.t("Choose the Microsoft-signed mdm.exe (7.0 or later) of your Visual Studio or Office installation"),
        filters: { "mdm.exe": ["exe"] }
    });

    if (!picked || picked.length === 0)
        return false;

    const file = picked[0].fsPath;

    if (!await isSignedByMicrosoft(file)) {
        vscode.window.showErrorMessage(vscode.l10n.t("{0} is not signed by Microsoft; it is not registered.", file));
        return false;
    }

    const run_ = vscode.l10n.t("Register");
    const answer = await vscode.window.showWarningMessage(
        vscode.l10n.t("Register {0} as the Machine Debug Manager? This runs it as administrator with /regserver and installs the Windows service \"Machine Debug Manager\" (it can be removed with mdm.exe /unregserver).", file),
        { modal: true }, run_);

    if (answer !== run_)
        return false;

    const code = await runElevated(file, ["/regserver"]);

    if (code !== 0) {
        vscode.window.showErrorMessage(vscode.l10n.t("Registering failed or was cancelled (exit code {0}).", code));
        return false;
    }

    vscode.window.showInformationMessage(vscode.l10n.t("The Machine Debug Manager is registered. Start the application after this, so that it registers its scripts."));
    return true;
}


async function addToGroup(report: AttachReport): Promise<boolean> {
    const account = `${process.env.USERDOMAIN ?? "."}\\${userName()}`;
    const add = vscode.l10n.t("Add");
    const answer = await vscode.window.showWarningMessage(
        report.groupExists
            ? vscode.l10n.t("Add {0} to the local group \"{1}\"? This runs as administrator. Sign out of Windows and in again afterwards.", account, GROUP)
            : vscode.l10n.t("Create the local group \"{1}\" and add {0} to it? This runs as administrator. Sign out of Windows and in again afterwards.", account, GROUP),
        { modal: true }, add);

    if (answer !== add)
        return false;

    const q = (text: string) => "'" + text.replace(/'/g, "''") + "'";
    const code = await runElevatedScript(
        `if (-not (Get-LocalGroup -Name ${q(GROUP)} -ErrorAction SilentlyContinue)) { New-LocalGroup -Name ${q(GROUP)} | Out-Null }; ` +
        `Add-LocalGroupMember -Group ${q(GROUP)} -Member ${q(account)} -ErrorAction Stop`);

    if (code !== 0) {
        vscode.window.showErrorMessage(vscode.l10n.t("Adding to the group failed or was cancelled (exit code {0}).", code));
        return false;
    }

    vscode.window.showInformationMessage(vscode.l10n.t("{0} was added to \"{1}\". Sign out of Windows and in again for it to take effect.", account, GROUP));
    return true;
}


// ============================================================
// the command and the check at start-up
// ============================================================

let channel: vscode.OutputChannel | undefined;

function writeReport(report: AttachReport): void {
    channel ??= vscode.window.createOutputChannel(vscode.l10n.t("VBScript Attach Debugging"));
    channel.clear();
    channel.appendLine(vscode.l10n.t("Environment check for attach debugging"));
    channel.appendLine("");

    for (const check of report.checks)
        channel.appendLine(`${check.ok ? "[ OK ]" : check.optional ? "[ -- ]" : "[FAIL]"} ${check.title}\n         ${check.detail}`);

    channel.appendLine("");
    channel.appendLine(report.ready
        ? vscode.l10n.t("Ready. Start the application after the Machine Debug Manager is registered, then use a configuration with \"request\": \"attach\".")
        : vscode.l10n.t("Not ready. See docs/DEBUGGER.md, \"Attaching to another application\"."));
}


/** The command "VBScript: Check Attach Debugging Environment". */
export async function showAttachCheck(): Promise<void> {
    if (process.platform !== "win32") {
        vscode.window.showInformationMessage(vscode.l10n.t("Attach debugging needs Windows."));
        return;
    }

    const report = await vscode.window.withProgress(
        { location: vscode.ProgressLocation.Notification, title: vscode.l10n.t("Checking the attach debugging environment...") },
        () => checkAttachEnvironment());

    writeReport(report);
    channel?.show(true);

    if (report.ready) {
        vscode.window.showInformationMessage(vscode.l10n.t("Attach debugging: the environment is ready."));
        return;
    }

    const problems = report.checks.filter(c => !c.ok && !c.optional).map(c => c.title).join("; ");
    const register = vscode.l10n.t("Register mdm.exe...");
    const group = vscode.l10n.t("Add me to the group...");
    const buttons: string[] = [];

    if (!report.mdmRegistered)
        buttons.push(register);

    if (!report.inGroup && !report.addedToGroup)
        buttons.push(group);

    const choice = await vscode.window.showWarningMessage(
        vscode.l10n.t("Attach debugging is not ready: {0}", problems), ...buttons);

    if ((choice === register && await registerMdm()) || (choice === group && await addToGroup(report)))
        await showAttachCheck();
}


/** At start-up: a notification when attach debugging cannot work yet (it can be turned off in the settings). */
export async function checkAttachOnStartup(): Promise<void> {
    if (process.platform !== "win32")
        return;

    const settings = vscode.workspace.getConfiguration("vbscript");

    if (!settings.get<boolean>("attach.checkOnStartup", true))
        return;

    const report = await checkAttachEnvironment();

    if (report.ready)
        return;

    const problems = report.checks.filter(c => !c.ok && !c.optional).map(c => c.title).join("; ");
    const details = vscode.l10n.t("Details");
    const never = vscode.l10n.t("Don't show again");
    const choice = await vscode.window.showWarningMessage(
        vscode.l10n.t("Attach debugging (to applications such as Altium DXP) is not ready: {0}", problems), details, never);

    if (choice === details)
        await showAttachCheck();
    else if (choice === never)
        await settings.update("attach.checkOnStartup", false, vscode.ConfigurationTarget.Global);
}
