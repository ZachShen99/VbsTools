import * as os from "os";
import * as path from "path";
import * as vscode from "vscode";

/**
 * Commands behind the buttons of the "Get Started with VbsTools" walkthrough
 * (contributes.walkthroughs in package.json): a sample script to try completion,
 * running and debugging on.
 */
export class Walkthrough {
    private sample: vscode.Uri | undefined;

    public register(): vscode.Disposable[] {
        return [
            vscode.commands.registerCommand("vbscript.walkthrough.newScript", () => this.newScript()),
            vscode.commands.registerCommand("vbscript.walkthrough.debug", () => this.debug())
        ];
    }


    /** Saves the sample script where the user chooses (Documents\hello.vbs by default) and opens it. */
    private async newScript(): Promise<void> {
        const documents = path.join(os.homedir(), "Documents");
        const target = await vscode.window.showSaveDialog({
            title: vscode.l10n.t("Save the sample script"),
            defaultUri: vscode.Uri.file(path.join(documents, "hello.vbs")),
            filters: { VBScript: ["vbs"] }
        });

        if (!target) {
            return;
        }

        const text = sampleScript();

        // Non-ASCII comments (translations): UTF-16 LE with BOM, which cscript runs too.
        const bytes = /[^\x00-\x7F]/.test(text)
            ? Buffer.concat([Buffer.from([0xFF, 0xFE]), Buffer.from(text, "utf16le")])
            : Buffer.from(text, "latin1");

        await vscode.workspace.fs.writeFile(target, bytes);
        this.sample = target;

        // Beside the walkthrough, so both stay visible.
        const document = await vscode.workspace.openTextDocument(target);
        const editor = await vscode.window.showTextDocument(document, { viewColumn: vscode.ViewColumn.Beside });

        // Cursor on the empty line below the instruction: typing "fso." shows the members.
        const line = document.getText().split(/\r?\n/).findIndex(l => l.startsWith("Set fso = "));

        if (line >= 0) {
            const position = new vscode.Position(line + 2, 0);
            editor.selection = new vscode.Selection(position, position);
        }
    }


    /** Starts debugging the sample (or the open script), with a breakpoint in its loop. */
    private async debug(): Promise<void> {
        const editor = scriptEditor();
        const uri = editor?.document.uri ?? this.sample;

        if (!uri) {
            void vscode.window.showInformationMessage(vscode.l10n.t("Open a VBScript file first."));
            return;
        }

        const document = await vscode.workspace.openTextDocument(uri);

        if (document.isDirty) {
            await document.save();
        }

        // A breakpoint on "count = count + 1" unless the script already has one.
        const hasBreakpoint = vscode.debug.breakpoints.some(b =>
            b instanceof vscode.SourceBreakpoint && b.location.uri.toString() === uri.toString());
        const line = document.getText().split(/\r?\n/).findIndex(l => /^\s*count = count \+ 1/.test(l));

        if (!hasBreakpoint && line >= 0) {
            vscode.debug.addBreakpoints([new vscode.SourceBreakpoint(new vscode.Location(uri, new vscode.Position(line, 0)))]);
        }

        await vscode.debug.startDebugging(undefined, {
            type: "vbscript",
            request: "launch",
            name: "Debug VBScript",
            program: uri.fsPath
        });
    }
}


/** The active VBScript editor, else a visible one (a walkthrough or other page may have the focus). */
export function scriptEditor(): vscode.TextEditor | undefined {
    const active = vscode.window.activeTextEditor;

    return active?.document.languageId === "vbscript"
        ? active
        : vscode.window.visibleTextEditors.find(e => e.document.languageId === "vbscript");
}


function sampleScript(): string {
    return [
        "Option Explicit",
        "",
        "' " + vscode.l10n.t("VbsTools sample script"),
        "",
        "Dim fso, folder, file, count",
        "Set fso = CreateObject(\"Scripting.FileSystemObject\")",
        "' " + vscode.l10n.t("Type fso. on the empty line below to see the members of FileSystemObject (then delete it again):"),
        "",
        "Set folder = fso.GetSpecialFolder(0)   ' " + vscode.l10n.t("the Windows folder"),
        "",
        "count = 0",
        "",
        "For Each file In folder.Files",
        "    count = count + 1                ' " + vscode.l10n.t("F9 sets a breakpoint here, F5 starts debugging"),
        "Next",
        "",
        "WScript.Echo folder.Path & \": \" & count & \" files\"",
        ""
    ].join("\r\n");
}
