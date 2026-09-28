import { execFileSync } from "child_process";
import * as vscode from "vscode";

/**
 * File encodings of .vbs files.
 *
 * Windows Script Host (cscript/wscript) reads a script as UTF-16 LE when it
 * starts with a BOM and otherwise with the system ANSI code page (GetACP);
 * it does not understand UTF-8. VS Code, on the other hand, opens and saves
 * files as UTF-8 by default. So:
 *
 *  - an ANSI script with Chinese/Japanese/accented text opens garbled in VS Code
 *    (and saving it would destroy the text) — reopened here with the ANSI code page;
 *  - a script saved as UTF-8 runs correctly in VbsTools Script Host (which detects
 *    UTF-8) but prints garbled text in cscript/wscript — reported as a warning with
 *    quick fixes that convert it to UTF-16 LE or ANSI.
 *
 * The detection matches VbScript.Hosting.ScriptEncoding (script host, debugger).
 */

interface CodePage {
    readonly codePage: number;
    /** VS Code encoding id (TextDocument.encoding). */
    readonly id: string;
    readonly name: string;
    /** Other VS Code ids that decode the same text. */
    readonly aliases?: readonly string[];
}

const CodePages: readonly CodePage[] = [
    { codePage: 936, id: "gbk", name: "GBK / 936", aliases: ["gb18030", "gb2312"] },
    { codePage: 932, id: "shiftjis", name: "Shift JIS / 932" },
    { codePage: 949, id: "euckr", name: "EUC-KR / 949" },
    { codePage: 950, id: "cp950", name: "Big5 / 950", aliases: ["big5hkscs"] },
    { codePage: 874, id: "windows874", name: "Windows-874" },
    { codePage: 1250, id: "windows1250", name: "Windows-1250" },
    { codePage: 1251, id: "windows1251", name: "Windows-1251" },
    { codePage: 1252, id: "windows1252", name: "Windows-1252", aliases: ["iso88591"] },
    { codePage: 1253, id: "windows1253", name: "Windows-1253" },
    { codePage: 1254, id: "windows1254", name: "Windows-1254" },
    { codePage: 1255, id: "windows1255", name: "Windows-1255" },
    { codePage: 1256, id: "windows1256", name: "Windows-1256" },
    { codePage: 1257, id: "windows1257", name: "Windows-1257" },
    { codePage: 1258, id: "windows1258", name: "Windows-1258" },
    // "Beta: Use Unicode UTF-8 for worldwide language support": WSH then reads UTF-8.
    { codePage: 65001, id: "utf8", name: "UTF-8" }
];

const Western = CodePages.find(c => c.codePage === 1252)!;

const DiagnosticCode = "encoding";


let systemCodePage: number | undefined;

/** The system ANSI code page (what cscript uses), from the registry; 1252 elsewhere. */
function systemAnsiCodePage(): number {
    if (systemCodePage !== undefined) {
        return systemCodePage;
    }

    systemCodePage = 1252;

    if (process.platform === "win32") {
        try {
            const output = execFileSync("reg",
                ["query", "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Nls\\CodePage", "/v", "ACP"],
                { encoding: "latin1", windowsHide: true, timeout: 5000 });
            const match = /ACP\s+REG_SZ\s+(\d+)/i.exec(output);

            if (match) {
                systemCodePage = parseInt(match[1], 10);
            }
        } catch {
            // keep 1252
        }
    }

    return systemCodePage;
}


/** The code page .vbs files without BOM are read with: vbscript.encoding.ansi, else the system's. */
export function ansiCodePage(): CodePage {
    const setting = vscode.workspace.getConfiguration("vbscript").get<string>("encoding.ansi", "auto");
    const chosen = CodePages.find(c => c.id === setting);

    if (chosen) {
        return chosen;
    }

    const system = systemAnsiCodePage();
    return CodePages.find(c => c.codePage === system) ?? Western;
}


function sameEncoding(id: string, codePage: CodePage): boolean {
    return id === codePage.id || (codePage.aliases?.includes(id) ?? false);
}


// ============================================================
// Detection
// ============================================================

const StrictUtf8 = new TextDecoder("utf-8", { fatal: true });

/** The encoding the file on disk is in, as the script host decides it (undefined: plain ASCII). */
export function detect(bytes: Uint8Array): "utf16le" | "utf8bom" | "utf8" | "ansi" | undefined {
    if (bytes.length >= 2 && bytes[0] === 0xFF && bytes[1] === 0xFE) {
        return "utf16le";
    }

    if (bytes.length >= 3 && bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) {
        return "utf8bom";
    }

    if (!bytes.some(b => b >= 0x80)) {
        return undefined;
    }

    try {
        StrictUtf8.decode(bytes);
        return "utf8";
    } catch {
        return "ansi";
    }
}


function firstNonAscii(document: vscode.TextDocument): vscode.Position | undefined {
    const text = document.getText();
    const index = text.search(/[^\x00-\x7F]/);

    return index < 0 ? undefined : document.positionAt(index);
}


interface Problem {
    readonly range: vscode.Range;
    readonly message: string;
}

/** Why cscript/wscript would not read this document as it appears in the editor. */
export function findProblem(document: vscode.TextDocument): Problem | undefined {
    const encoding = document.encoding;
    const ansi = ansiCodePage();
    const lineOne = document.lineAt(0).range;

    switch (encoding) {
        case "utf16le":
            return undefined;

        case "utf8bom":
            return {
                range: lineOne,
                message: vscode.l10n.t("Encoded as UTF-8 with BOM: cscript and wscript cannot run it (\"Invalid character\" on line 1). Save it as UTF-16 LE or ANSI.")
            };

        case "utf16be":
            return {
                range: lineOne,
                message: vscode.l10n.t("Encoded as UTF-16 BE: cscript and wscript cannot run it. Save it as UTF-16 LE.")
            };
    }

    const position = firstNonAscii(document);

    if (!position || sameEncoding(encoding, ansi)) {
        return undefined;
    }

    const range = new vscode.Range(position, position.translate(0, 1));

    return encoding === "utf8"
        ? {
            range,
            message: vscode.l10n.t("Encoded as UTF-8: cscript and wscript read the file as ANSI ({0}) and show text that is not ASCII garbled (VbsTools Script Host reads it correctly). Save it as UTF-16 LE or ANSI.", ansi.name)
        }
        : {
            range,
            message: vscode.l10n.t("Encoded as {0}: cscript and wscript read the file as ANSI ({1}) and show text that is not ASCII garbled.", encoding, ansi.name)
        };
}


// ============================================================
// Converting
// ============================================================

/** Rewrites the file in another encoding and reopens it with that encoding. */
async function saveWithEncoding(document: vscode.TextDocument, target: string): Promise<boolean> {
    if (document.isUntitled) {
        void vscode.window.showInformationMessage(vscode.l10n.t("Save the script first."));
        return false;
    }

    if (document.isDirty && !(await document.save())) {
        return false;
    }

    const text = document.getText();
    let bytes: Uint8Array;

    if (target === "utf16le") {
        bytes = Buffer.concat([Buffer.from([0xFF, 0xFE]), Buffer.from(text, "utf16le")]);
    } else if (target === "utf8") {
        bytes = Buffer.from(text, "utf8");
    } else {
        bytes = await vscode.workspace.encode(text, { encoding: target });

        if (await vscode.workspace.decode(bytes, { encoding: target }) !== text) {
            void vscode.window.showErrorMessage(vscode.l10n.t("The script contains characters that {0} cannot represent. Save it as UTF-16 LE instead.", target));
            return false;
        }
    }

    await vscode.workspace.fs.writeFile(document.uri, bytes);

    // Decode the new bytes with the new encoding. VS Code may reload the file on its
    // own at the same time (with the old encoding): check and repeat once.
    for (let attempt = 0; attempt < 3; attempt++) {
        const reopened = await vscode.workspace.openTextDocument(document.uri, { encoding: target });

        if (reopened.getText() === text && sameId(reopened.encoding, target)) {
            return true;
        }

        await new Promise(resolve => setTimeout(resolve, 300));
    }

    void vscode.window.showWarningMessage(vscode.l10n.t("The file was saved as {0}; reopen it to see the text.", target));
    return true;
}


function sameId(actual: string, target: string): boolean {
    return actual === target || CodePages.some(c => c.id === target && sameEncoding(actual, c));
}


interface EncodingPick extends vscode.QuickPickItem {
    readonly id: string;
}

async function pickEncoding(document: vscode.TextDocument): Promise<string | undefined> {
    const ansi = ansiCodePage();
    const items: EncodingPick[] = [
        {
            id: "utf16le",
            label: "UTF-16 LE",
            description: vscode.l10n.t("cscript, wscript and VbsTools — any character"),
            picked: document.encoding === "utf16le"
        }
    ];

    if (ansi.codePage !== 65001) {
        items.push({
            id: ansi.id,
            label: `ANSI — ${ansi.name}`,
            description: vscode.l10n.t("cscript, wscript and VbsTools on systems with code page {0}", ansi.codePage),
            picked: sameEncoding(document.encoding, ansi)
        });
    }

    items.push({
        id: "utf8",
        label: "UTF-8",
        description: ansi.codePage === 65001
            ? vscode.l10n.t("cscript, wscript and VbsTools")
            : vscode.l10n.t("VbsTools Script Host only (cscript reads text that is not ASCII wrong)"),
        picked: document.encoding === "utf8"
    });

    const pick = await vscode.window.showQuickPick(items, {
        placeHolder: vscode.l10n.t("Save the script with encoding (currently {0})", document.encoding)
    });

    return pick?.id;
}


// ============================================================
// Registration
// ============================================================

export class ScriptEncodings implements vscode.CodeActionProvider {
    private readonly diagnostics = vscode.languages.createDiagnosticCollection("vbscript-encoding");

    /** Documents reopened with the ANSI code page once (never fight a later manual choice). */
    private readonly reopened = new Set<string>();

    public register(): vscode.Disposable[] {
        const disposables: vscode.Disposable[] = [
            this.diagnostics,
            vscode.commands.registerCommand("vbscript.saveWithEncoding", async (uri?: vscode.Uri, target?: string) => {
                const document = uri
                    ? await vscode.workspace.openTextDocument(uri)
                    : vscode.window.activeTextEditor?.document;

                if (!document || document.languageId !== "vbscript") {
                    return;
                }

                const encoding = target ?? await pickEncoding(document);

                if (encoding) {
                    await saveWithEncoding(document, encoding);
                }
            }),
            vscode.languages.registerCodeActionsProvider({ language: "vbscript", scheme: "file" }, this,
                { providedCodeActionKinds: [vscode.CodeActionKind.QuickFix] }),
            vscode.workspace.onDidOpenTextDocument(document => void this.opened(document)),
            vscode.workspace.onDidChangeTextDocument(e => this.check(e.document)),
            vscode.workspace.onDidCloseTextDocument(document => {
                this.diagnostics.delete(document.uri);
                this.reopened.delete(document.uri.toString());
            }),
            vscode.workspace.onDidChangeConfiguration(e => {
                if (e.affectsConfiguration("vbscript.encoding")) {
                    vscode.workspace.textDocuments.forEach(d => this.check(d));
                }
            })
        ];

        vscode.workspace.textDocuments.forEach(d => void this.opened(d));
        return disposables;
    }


    private async opened(document: vscode.TextDocument): Promise<void> {
        if (!isScriptFile(document)) {
            return;
        }

        await this.reopenAnsi(document);
        this.check(document);
    }


    /**
     * VS Code decoded an ANSI script as UTF-8 (the text shows replacement
     * characters): reopen it with the ANSI code page, as cscript reads it.
     * Only when VS Code used plain UTF-8 — an encoding VS Code guessed or the
     * user picked is left alone.
     */
    private async reopenAnsi(document: vscode.TextDocument): Promise<void> {
        const key = document.uri.toString();

        if (document.isDirty || document.encoding !== "utf8" || this.reopened.has(key)) {
            return;
        }

        const ansi = ansiCodePage();

        if (ansi.codePage === 65001) {
            return;
        }

        let bytes: Uint8Array;

        try {
            bytes = await vscode.workspace.fs.readFile(document.uri);
        } catch {
            return;
        }

        if (detect(bytes) !== "ansi") {
            return;
        }

        this.reopened.add(key);

        try {
            await vscode.workspace.openTextDocument(document.uri, { encoding: ansi.id });
        } catch {
            // edited in the meantime: leave it
        }
    }


    private check(document: vscode.TextDocument): void {
        if (!isScriptFile(document)) {
            return;
        }

        const problem = vscode.workspace.getConfiguration("vbscript").get<boolean>("encoding.check", true)
            ? findProblem(document)
            : undefined;

        if (!problem) {
            this.diagnostics.delete(document.uri);
            return;
        }

        const diagnostic = new vscode.Diagnostic(problem.range, problem.message, vscode.DiagnosticSeverity.Warning);
        diagnostic.source = "vbscript";
        diagnostic.code = DiagnosticCode;
        this.diagnostics.set(document.uri, [diagnostic]);
    }


    public async provideCodeActions(document: vscode.TextDocument, _range: vscode.Range,
        context: vscode.CodeActionContext): Promise<vscode.CodeAction[]> {
        const diagnostic = context.diagnostics.find(d => d.code === DiagnosticCode);

        if (!diagnostic) {
            return [];
        }

        const actions: vscode.CodeAction[] = [];
        const add = (title: string, target: string, preferred: boolean) => {
            const action = new vscode.CodeAction(title, vscode.CodeActionKind.QuickFix);
            action.diagnostics = [diagnostic];
            action.isPreferred = preferred;
            action.command = { title, command: "vbscript.saveWithEncoding", arguments: [document.uri, target] };
            actions.push(action);
        };

        const ansi = ansiCodePage();
        const text = document.getText();
        const fitsAnsi = ansi.codePage !== 65001 &&
            await vscode.workspace.decode(await vscode.workspace.encode(text, { encoding: ansi.id }), { encoding: ansi.id }) === text;

        // UTF-16 LE first (as VbsTools Studio saves): it runs on any system; ANSI only
        // on systems with the same code page, but keeps diffs readable.
        add(vscode.l10n.t("Save as UTF-16 LE (runs in cscript and wscript)"), "utf16le", true);

        if (fitsAnsi) {
            add(vscode.l10n.t("Save as ANSI ({0})", ansi.name), ansi.id, false);
        }

        return actions;
    }
}


function isScriptFile(document: vscode.TextDocument): boolean {
    return document.languageId === "vbscript" && document.uri.scheme === "file";
}
