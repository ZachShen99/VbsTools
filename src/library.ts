import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import * as vscode from "vscode";

/**
 * The "Code Library" view (VbsTools activity bar):
 * - Samples: the sample collection of VbsEdit (samples.dat / samples2.dat, read from the
 *   user's VbsEdit installation or "vbscript.library.samplesFile"), else VbsTools Studio's samples;
 * - My Snippets: the user's snippets (%APPDATA%\VbsTools\snippets.json, shared with VbsTools
 *   Studio), each with a shortcut that expands it while typing; edited in place (Ctrl+S saves).
 */

type Node = GroupNode | FolderNode | SampleNode | SnippetNode;

interface GroupNode { kind: "group"; id: "samples" | "mine"; label: string; }
interface FolderNode { kind: "folder"; label: string; group: "samples" | "mine"; path: string; children: Node[]; }
interface SampleNode { kind: "sample"; label: string; text: () => string; }
interface SnippetNode { kind: "snippet"; label: string; snippet: Snippet; }

export interface Snippet { name: string; prefix: string; body: string; description: string; folder: string; }

export const PREVIEW_SCHEME = "vbstools-library";
export const SNIPPET_SCHEME = "vbstools-snippet";

const KEYWORDS = new Set(("and as byref byval call case class const default dim do each else elseif empty end eqv erase error exit explicit false for function get goto " +
    "if imp in is let loop me mod new next not nothing null on option or preserve private property public randomize redim rem resume select set step stop sub " +
    "then to true until wend while with xor").split(" "));


export class CodeLibrary implements vscode.TreeDataProvider<Node>, vscode.TextDocumentContentProvider {
    private readonly changed = new vscode.EventEmitter<Node | undefined>();
    readonly onDidChangeTreeData = this.changed.event;

    private readonly previewChanged = new vscode.EventEmitter<vscode.Uri>();
    readonly onDidChange = this.previewChanged.event;

    private filter = "";
    private readonly previews = new Map<string, string>();
    private view: vscode.TreeView<Node> | undefined;
    private samples: { source: string; mtime: number; root: FolderNode } | undefined;
    readonly store = new SnippetStore();

    constructor(private readonly context: vscode.ExtensionContext) { }


    register(): vscode.Disposable[] {
        this.view = vscode.window.createTreeView("vbscript.library", { treeDataProvider: this, showCollapseAll: true });
        this.store.onDidChange(() => this.changed.fire(undefined));

        const command = (id: string, handler: (...args: any[]) => unknown) => vscode.commands.registerCommand(id, handler);

        return [
            this.view,
            this.store,
            vscode.workspace.registerTextDocumentContentProvider(PREVIEW_SCHEME, this),
            vscode.workspace.registerFileSystemProvider(SNIPPET_SCHEME, new SnippetFileSystem(this.store), { isCaseSensitive: false }),
            vscode.workspace.onDidChangeConfiguration(e => {
                if (e.affectsConfiguration("vbscript.library")) {
                    this.samples = undefined;
                    this.changed.fire(undefined);
                }
            }),
            command("vbscript.library.refresh", () => { this.samples = undefined; this.store.reload(); this.changed.fire(undefined); }),
            command("vbscript.library.search", () => this.search()),
            command("vbscript.library.clearSearch", () => this.setFilter("")),
            command("vbscript.library.preview", (node: Node) => this.preview(node)),
            command("vbscript.library.insert", (node: Node) => this.insert(node)),
            command("vbscript.library.newFile", (node: Node) => this.newFile(node)),
            command("vbscript.library.copyToSnippets", (node: Node) => this.copyToSnippets(node)),
            command("vbscript.library.newSnippet", (node?: Node) => this.newSnippet(node)),
            command("vbscript.library.addSelection", () => this.addSelection()),
            command("vbscript.library.editSnippet", (node: Node) => this.editSnippet(node)),
            command("vbscript.library.changeShortcut", (node: Node) => this.changeShortcut(node)),
            command("vbscript.library.renameSnippet", (node: Node) => this.renameSnippet(node)),
            command("vbscript.library.moveSnippet", (node: Node) => this.moveSnippet(node)),
            command("vbscript.library.deleteSnippet", (node: Node) => this.deleteSnippet(node)),
            command("vbscript.library.openSnippetsFile", () => this.openSnippetsFile()),
            vscode.languages.registerCompletionItemProvider({ language: "vbscript" }, new UserSnippetCompletion(this.store))
        ];
    }


    // ============================================================
    // Tree
    // ============================================================

    getTreeItem(node: Node): vscode.TreeItem {
        switch (node.kind) {
            case "group": {
                const item = new vscode.TreeItem(node.label, vscode.TreeItemCollapsibleState.Expanded);
                item.contextValue = "group-" + node.id;
                item.iconPath = new vscode.ThemeIcon(node.id === "samples" ? "book" : "symbol-snippet");

                if (node.id === "samples") {
                    item.tooltip = this.sampleRoot().path;
                } else {
                    item.tooltip = this.store.file;
                    item.description = `${this.store.all().length}`;
                }

                return item;
            }

            case "folder": {
                const item = new vscode.TreeItem(node.label,
                    this.filter ? vscode.TreeItemCollapsibleState.Expanded : vscode.TreeItemCollapsibleState.Collapsed);
                item.contextValue = node.group === "mine" ? "folder-mine" : "folder";
                item.iconPath = vscode.ThemeIcon.Folder;
                return item;
            }

            case "sample": {
                const item = new vscode.TreeItem(node.label, vscode.TreeItemCollapsibleState.None);
                item.contextValue = "sample";
                item.iconPath = new vscode.ThemeIcon("file-code");
                item.command = { command: "vbscript.library.preview", title: vscode.l10n.t("Preview"), arguments: [node] };
                return item;
            }

            case "snippet": {
                const s = node.snippet;
                const item = new vscode.TreeItem(node.label, vscode.TreeItemCollapsibleState.None);
                item.description = s.prefix;
                item.contextValue = "snippet";
                item.iconPath = new vscode.ThemeIcon("symbol-snippet");
                item.tooltip = new vscode.MarkdownString(`**${s.name}** — ${vscode.l10n.t("shortcut")} \`${s.prefix}\`${s.description ? "\n\n" + s.description : ""}\n\n`)
                    .appendCodeblock(firstLines(snippetText(s.body), 15), "vbscript");
                item.command = { command: "vbscript.library.preview", title: vscode.l10n.t("Preview"), arguments: [node] };
                return item;
            }
        }
    }


    getChildren(node?: Node): Node[] {
        if (!node) {
            const groups: GroupNode[] = [
                { kind: "group", id: "samples", label: vscode.l10n.t("Samples") },
                { kind: "group", id: "mine", label: vscode.l10n.t("My Snippets") }
            ];
            return this.filter ? groups.filter(g => this.getChildren(g).length > 0) : groups;
        }

        switch (node.kind) {
            case "group":
                return this.filtered(node.id === "samples" ? this.sampleRoot().children : this.snippetTree().children);

            case "folder":
                return this.filtered(node.children);

            default:
                return [];
        }
    }


    private filtered(nodes: Node[]): Node[] {
        if (!this.filter) {
            return nodes;
        }

        const needle = this.filter.toLowerCase();
        const matches = (n: Node): boolean => {
            switch (n.kind) {
                case "folder": return n.children.some(matches);
                case "sample": return n.label.toLowerCase().includes(needle) || n.text().toLowerCase().includes(needle);
                case "snippet": return [n.snippet.name, n.snippet.prefix, n.snippet.description, n.snippet.body].some(t => t.toLowerCase().includes(needle));
                default: return true;
            }
        };

        return nodes.filter(matches);
    }


    /** Samples: VbsEdit's samples file when found, else VbsTools Studio's samples folder. */
    private sampleRoot(): FolderNode {
        const file = samplesDataFile();

        if (file) {
            const mtime = fs.statSync(file).mtimeMs;

            if (!this.samples || this.samples.source !== file || this.samples.mtime !== mtime) {
                try {
                    this.samples = { source: file, mtime, root: readSamplesData(file) };
                } catch (error) {
                    void vscode.window.showWarningMessage(vscode.l10n.t("Cannot read the samples file {0}: {1}", file, String(error)));
                    this.samples = undefined;
                }
            }

            if (this.samples) {
                return this.samples.root;
            }
        }

        const folder = builtinSamplesFolder(this.context);
        return { kind: "folder", label: "Samples", group: "samples", path: folder ?? "", children: folder ? readSamplesFolder(folder) : [] };
    }


    private snippetTree(): FolderNode {
        const root: FolderNode = { kind: "folder", label: "", group: "mine", path: "", children: [] };

        for (const snippet of this.store.all()) {
            let parent = root;

            for (const part of snippet.folder.split("/").map(p => p.trim()).filter(p => p)) {
                let child = parent.children.find((c): c is FolderNode => c.kind === "folder" && c.label.toLowerCase() === part.toLowerCase());

                if (!child) {
                    child = { kind: "folder", label: part, group: "mine", path: parent.path ? `${parent.path}/${part}` : part, children: [] };
                    parent.children.push(child);
                }

                parent = child;
            }

            parent.children.push({ kind: "snippet", label: snippet.name, snippet });
        }

        const sort = (folder: FolderNode): void => {
            folder.children.sort((a, b) => (a.kind === "folder" ? 0 : 1) - (b.kind === "folder" ? 0 : 1) || a.label.localeCompare(b.label, undefined, { numeric: true }));
            folder.children.forEach(c => c.kind === "folder" && sort(c));
        };

        sort(root);
        return root;
    }


    // ============================================================
    // Preview, insert, new file
    // ============================================================

    private async search(): Promise<void> {
        const text = await vscode.window.showInputBox({
            title: vscode.l10n.t("Search the Code Library"),
            prompt: vscode.l10n.t("Names and contents of samples and snippets"),
            value: this.filter
        });

        if (text !== undefined) {
            this.setFilter(text.trim());
        }
    }


    private setFilter(filter: string): void {
        this.filter = filter;
        void vscode.commands.executeCommand("setContext", "vbscript.librarySearch", filter.length > 0);

        if (this.view) {
            this.view.message = filter ? vscode.l10n.t("Search: {0}", filter) : undefined;
        }

        this.changed.fire(undefined);
    }


    provideTextDocumentContent(uri: vscode.Uri): string {
        return this.previews.get(uri.toString()) ?? "";
    }


    private async preview(node: Node): Promise<void> {
        const text = contentOf(node);

        if (text === undefined) {
            return;
        }

        const uri = vscode.Uri.from({ scheme: PREVIEW_SCHEME, path: "/" + safeName(node.label) + ".vbs" });
        this.previews.set(uri.toString(), text);
        this.previewChanged.fire(uri);

        const document = await vscode.workspace.openTextDocument(uri);
        await vscode.languages.setTextDocumentLanguage(document, "vbscript");
        await vscode.window.showTextDocument(document, { preview: true, preserveFocus: true });
    }


    private async insert(node: Node): Promise<void> {
        const editor = targetEditor();

        if (!editor) {
            void vscode.window.showInformationMessage(vscode.l10n.t("Open a script to insert into."));
            return;
        }

        if (node.kind === "snippet") {
            await editor.insertSnippet(new vscode.SnippetString(node.snippet.body));
        } else if (node.kind === "sample") {
            const text = node.text();
            await editor.edit(edit => editor.selections.forEach(s => edit.replace(s, text)));
        } else {
            return;
        }

        await vscode.window.showTextDocument(editor.document, editor.viewColumn);
    }


    private async newFile(node: Node): Promise<void> {
        const text = contentOf(node);

        if (text === undefined) {
            return;
        }

        const document = await vscode.workspace.openTextDocument({ language: "vbscript", content: "" });
        const editor = await vscode.window.showTextDocument(document);

        if (node.kind === "snippet") {
            await editor.insertSnippet(new vscode.SnippetString(node.snippet.body));
        } else {
            await editor.edit(edit => edit.insert(new vscode.Position(0, 0), text));
        }
    }


    // ============================================================
    // My Snippets
    // ============================================================

    /** Editor context menu: the selection (or the whole script) becomes a snippet. */
    private async addSelection(): Promise<void> {
        const editor = vscode.window.activeTextEditor;

        if (!editor) {
            return;
        }

        const text = editor.selection.isEmpty ? editor.document.getText() : editor.document.getText(editor.selection);

        if (text.trim()) {
            await this.createSnippet(escapeSnippet(text), suggestName(text), "");
        }
    }


    private async copyToSnippets(node: Node): Promise<void> {
        if (node.kind === "sample") {
            await this.createSnippet(escapeSnippet(node.text()), node.label, "");
        }
    }


    /** "New Snippet" (view title or a folder of My Snippets): opens the new snippet for editing. */
    private async newSnippet(node?: Node): Promise<void> {
        const folder = node?.kind === "folder" && node.group === "mine" ? node.path : "";
        await this.createSnippet("", "", folder, true);
    }


    private async createSnippet(body: string, suggested: string, folder: string, edit = false): Promise<void> {
        const name = await vscode.window.showInputBox({
            title: vscode.l10n.t("New snippet (1/3): name"),
            value: suggested,
            validateInput: v => this.store.validateName(v.trim())
        });

        if (!name) {
            return;
        }

        const prefix = await this.askShortcut(suggestPrefix(name.trim(), this.store), undefined, vscode.l10n.t("New snippet (2/3): shortcut"));

        if (!prefix) {
            return;
        }

        const target = folder || await this.askFolder(vscode.l10n.t("New snippet (3/3): folder"));

        if (target === undefined) {
            return;
        }

        const snippet: Snippet = { name: name.trim(), prefix, body, description: "", folder: target };
        this.store.add(snippet);

        if (edit || !body) {
            await this.openSnippetEditor(snippet);
        } else {
            void vscode.window.showInformationMessage(vscode.l10n.t("Snippet \"{0}\" added: type {1} to insert it.", snippet.name, prefix));
        }
    }


    /** Shortcut that expands the snippet; must not clash with keywords or other snippets. */
    private askShortcut(value: string, current: Snippet | undefined, title: string): Thenable<string | undefined> {
        return vscode.window.showInputBox({
            title,
            prompt: vscode.l10n.t("Typed in a script, then chosen from the suggestions (or Tab), it inserts the snippet"),
            value,
            validateInput: v => this.store.validatePrefix(v.trim(), current, builtinPrefixes(this.context))
        }).then(v => v?.trim());
    }


    private async askFolder(title: string, current = ""): Promise<string | undefined> {
        const folders = this.store.folders();
        const picked = await vscode.window.showQuickPick(
            [
                { label: "$(root-folder) " + vscode.l10n.t("(top level)"), folder: "" },
                ...folders.map(f => ({ label: "$(folder) " + f, folder: f, picked: f === current })),
                { label: "$(new-folder) " + vscode.l10n.t("New folder…"), folder: "\0" }
            ],
            { title });

        if (!picked) {
            return undefined;
        }

        if (picked.folder !== "\0") {
            return picked.folder;
        }

        const created = await vscode.window.showInputBox({
            title: vscode.l10n.t("New folder (use / for subfolders)"),
            validateInput: v => /^[^\\:*?"<>|]+$/.test(v.trim()) ? undefined : vscode.l10n.t("Not a valid folder name")
        });

        return created?.trim().replace(/^\/+|\/+$/g, "") || undefined;
    }


    private async editSnippet(node: Node): Promise<void> {
        if (node.kind === "snippet") {
            await this.openSnippetEditor(node.snippet);
        }
    }


    /** The body in an editor: Ctrl+S writes it back to snippets.json. */
    private async openSnippetEditor(snippet: Snippet): Promise<void> {
        const document = await vscode.workspace.openTextDocument(snippetUri(snippet.name));
        await vscode.languages.setTextDocumentLanguage(document, "vbscript");
        await vscode.window.showTextDocument(document, { preview: false });
        void vscode.window.setStatusBarMessage(vscode.l10n.t("Editing snippet \"{0}\" (shortcut {1}) — Ctrl+S saves; {2} marks a placeholder", snippet.name, snippet.prefix, "${1:name}"), 8000);
    }


    private async changeShortcut(node: Node): Promise<void> {
        if (node.kind !== "snippet") {
            return;
        }

        const prefix = await this.askShortcut(node.snippet.prefix, node.snippet, vscode.l10n.t("Shortcut of \"{0}\"", node.snippet.name));

        if (prefix && prefix !== node.snippet.prefix) {
            this.store.update(node.snippet.name, { prefix });
        }
    }


    private async renameSnippet(node: Node): Promise<void> {
        if (node.kind !== "snippet") {
            return;
        }

        const name = await vscode.window.showInputBox({
            title: vscode.l10n.t("Rename \"{0}\"", node.snippet.name),
            value: node.snippet.name,
            validateInput: v => v.trim().toLowerCase() === node.snippet.name.toLowerCase() ? undefined : this.store.validateName(v.trim())
        });

        if (name && name.trim() !== node.snippet.name) {
            this.store.update(node.snippet.name, { name: name.trim() });
        }
    }


    private async moveSnippet(node: Node): Promise<void> {
        if (node.kind !== "snippet") {
            return;
        }

        const folder = await this.askFolder(vscode.l10n.t("Move \"{0}\" to", node.snippet.name), node.snippet.folder);

        if (folder !== undefined) {
            this.store.update(node.snippet.name, { folder });
        }
    }


    private async deleteSnippet(node: Node): Promise<void> {
        if (node.kind !== "snippet") {
            return;
        }

        const remove = vscode.l10n.t("Delete");
        const answer = await vscode.window.showWarningMessage(vscode.l10n.t("Delete the snippet \"{0}\"?", node.snippet.name), { modal: true }, remove);

        if (answer === remove) {
            this.store.remove(node.snippet.name);
        }
    }


    private async openSnippetsFile(): Promise<void> {
        this.store.ensureFile();
        await vscode.window.showTextDocument(vscode.Uri.file(this.store.file));
    }
}


// ============================================================
// User snippets (%APPDATA%\VbsTools\snippets.json)
// ============================================================

export class SnippetStore implements vscode.Disposable {
    readonly file = path.join(process.env.APPDATA ?? path.join(os.homedir(), "AppData", "Roaming"), "VbsTools", "snippets.json");

    private readonly changed = new vscode.EventEmitter<void>();
    readonly onDidChange = this.changed.event;

    private snippets: Snippet[] | undefined;
    private watcher: fs.FSWatcher | undefined;
    private writing = false;

    constructor() {
        this.watch();
    }


    all(): Snippet[] {
        this.snippets ??= readSnippets(this.file);
        return this.snippets;
    }

    find(name: string): Snippet | undefined {
        return this.all().find(s => s.name.toLowerCase() === name.toLowerCase());
    }

    folders(): string[] {
        return [...new Set(this.all().map(s => s.folder).filter(f => f))].sort((a, b) => a.localeCompare(b));
    }

    reload(): void {
        this.snippets = undefined;
        this.changed.fire();
    }


    validateName(name: string): string | undefined {
        if (!name) {
            return vscode.l10n.t("Enter a name");
        }

        return this.find(name) ? vscode.l10n.t("A snippet named \"{0}\" exists", name) : undefined;
    }


    /** Shortcuts must be words, not VBScript keywords, and not used by another snippet (built-in or yours). */
    validatePrefix(prefix: string, current: Snippet | undefined, builtin: Map<string, string>): string | undefined {
        if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(prefix)) {
            return vscode.l10n.t("Letters, digits and _ (starting with a letter)");
        }

        const key = prefix.toLowerCase();

        if (KEYWORDS.has(key)) {
            return vscode.l10n.t("\"{0}\" is a VBScript keyword", prefix);
        }

        if (builtin.has(key)) {
            return vscode.l10n.t("Used by the built-in snippet \"{0}\"", builtin.get(key) ?? "");
        }

        const other = this.all().find(s => s.prefix.toLowerCase() === key && s !== current);
        return other ? vscode.l10n.t("Used by your snippet \"{0}\"", other.name) : undefined;
    }


    add(snippet: Snippet): void {
        this.save([...this.all(), snippet]);
    }

    update(name: string, change: Partial<Snippet>): void {
        this.save(this.all().map(s => s.name.toLowerCase() === name.toLowerCase() ? { ...s, ...change } : s));
    }

    remove(name: string): void {
        this.save(this.all().filter(s => s.name.toLowerCase() !== name.toLowerCase()));
    }


    ensureFile(): void {
        if (!fs.existsSync(this.file)) {
            this.save(this.all());
        }
    }


    /** Same format as VbsTools Studio (VS Code snippet format + "folder"). */
    private save(snippets: Snippet[]): void {
        const json: Record<string, unknown> = {};

        for (const s of snippets) {
            json[s.name] = {
                prefix: s.prefix,
                body: s.body.split(/\r?\n/),
                ...(s.description ? { description: s.description } : {}),
                ...(s.folder ? { folder: s.folder } : {})
            };
        }

        fs.mkdirSync(path.dirname(this.file), { recursive: true });
        this.writing = true;

        try {
            fs.writeFileSync(this.file,
                "// VbsTools user snippets, shared by VbsTools Studio and the VS Code extension.\n" + JSON.stringify(json, null, "\t") + "\n", "utf8");
        } finally {
            setTimeout(() => { this.writing = false; }, 300);
        }

        this.snippets = snippets;
        this.watch();
        this.changed.fire();
    }


    /** Changes made by VbsTools Studio appear here. */
    private watch(): void {
        if (this.watcher || !fs.existsSync(this.file)) {
            return;
        }

        try {
            this.watcher = fs.watch(this.file, () => {
                if (!this.writing) {
                    this.reload();
                }
            });
        } catch {
            this.watcher = undefined;
        }
    }


    dispose(): void {
        this.watcher?.close();
    }
}


/** vbstools-snippet:/Name.vbs — a snippet's body as a document; saving writes it back. */
class SnippetFileSystem implements vscode.FileSystemProvider {
    private readonly changed = new vscode.EventEmitter<vscode.FileChangeEvent[]>();
    readonly onDidChangeFile = this.changed.event;

    constructor(private readonly store: SnippetStore) { }

    private snippet(uri: vscode.Uri): Snippet {
        const name = decodeURIComponent(path.posix.basename(uri.path).replace(/\.vbs$/i, ""));
        const snippet = this.store.find(name);

        if (!snippet) {
            throw vscode.FileSystemError.FileNotFound(uri);
        }

        return snippet;
    }

    watch(): vscode.Disposable { return new vscode.Disposable(() => undefined); }

    stat(uri: vscode.Uri): vscode.FileStat {
        const snippet = this.snippet(uri);
        return { type: vscode.FileType.File, ctime: 0, mtime: Date.now(), size: Buffer.byteLength(snippet.body) };
    }

    readDirectory(): [string, vscode.FileType][] { return []; }

    createDirectory(): void { throw vscode.FileSystemError.NoPermissions(); }

    readFile(uri: vscode.Uri): Uint8Array {
        return Buffer.from(this.snippet(uri).body.replace(/\r?\n/g, "\r\n"), "utf8");
    }

    writeFile(uri: vscode.Uri, content: Uint8Array): void {
        const snippet = this.snippet(uri);
        this.store.update(snippet.name, { body: Buffer.from(content).toString("utf8").replace(/\r\n/g, "\n") });
        this.changed.fire([{ type: vscode.FileChangeType.Changed, uri }]);
    }

    delete(): void { throw vscode.FileSystemError.NoPermissions(); }

    rename(): void { throw vscode.FileSystemError.NoPermissions(); }
}


/** User snippets complete by their shortcut (VS Code only knows the built-in ones from package.json). */
class UserSnippetCompletion implements vscode.CompletionItemProvider {
    constructor(private readonly store: SnippetStore) { }

    provideCompletionItems(document: vscode.TextDocument, position: vscode.Position, _token: vscode.CancellationToken, context: vscode.CompletionContext): vscode.CompletionItem[] {
        const before = document.lineAt(position.line).text.substring(0, position.character);

        if (context.triggerCharacter === "." || /\.\w*$/.test(before)) {
            return [];
        }

        return this.store.all().map(s => {
            const item = new vscode.CompletionItem({ label: s.prefix, description: s.name }, vscode.CompletionItemKind.Snippet);
            item.insertText = new vscode.SnippetString(s.body);
            item.documentation = new vscode.MarkdownString(s.description ? s.description + "\n\n" : "").appendCodeblock(snippetText(s.body), "vbscript");
            item.detail = vscode.l10n.t("My snippet");
            item.sortText = "!" + s.prefix;   // before words and COM members with the same letters
            return item;
        });
    }
}


// ============================================================
// Samples
// ============================================================

/** VbsEdit's samples file: the setting, or a VbsEdit installation. */
function samplesDataFile(): string | undefined {
    const configured = vscode.workspace.getConfiguration("vbscript").get<string>("library.samplesFile", "").trim();

    if (configured) {
        const file = configured.replace(/%([^%]+)%/g, (_, name: string) => process.env[name] ?? `%${name}%`);
        return fs.existsSync(file) ? file : undefined;
    }

    const roots = [process.env.ProgramFiles, process.env["ProgramFiles(x86)"], process.env.LOCALAPPDATA && path.join(process.env.LOCALAPPDATA, "Programs")]
        .filter((r): r is string => !!r);

    for (const root of roots) {
        for (const name of ["samples2.dat", "samples.dat"]) {
            const file = path.join(root, "Vbsedit", name);

            if (fs.existsSync(file)) {
                return file;
            }
        }
    }

    return undefined;
}


/**
 * VbsEdit's samples file (MFC archive): a folder is an item count, items
 * (int32 index, name, code), a subfolder count and subfolders (int32 index, name,
 * folder). Strings: length byte (0xFF: 16-bit length follows; 0xFFFF: 32-bit), ANSI text.
 */
export function readSamplesData(file: string): FolderNode {
    const data = fs.readFileSync(file);
    let position = 4;   // signature

    const int32 = (): number => {
        const value = data.readInt32LE(position);
        position += 4;
        return value;
    };

    const text = (): string => {
        let length = data[position++];

        if (length === 0xFF) {
            length = data.readUInt16LE(position);
            position += 2;

            if (length === 0xFFFF) {
                length = data.readUInt32LE(position);
                position += 4;
            }
        }

        const value = data.toString("latin1", position, position + length);
        position += length;
        return value;
    };

    const folder = (label: string, parentPath: string): FolderNode => {
        const node: FolderNode = { kind: "folder", label, group: "samples", path: parentPath ? `${parentPath}/${label}` : label, children: [] };
        const items = int32();

        for (let i = 0; i < items; i++) {
            int32();
            const name = text();
            const code = text();
            node.children.push({ kind: "sample", label: name, text: () => code });
        }

        const folders = int32();

        for (let i = 0; i < folders; i++) {
            int32();
            node.children.push(folder(text(), node.path));
        }

        return node;
    };

    const root = folder("Samples", "");
    root.path = file;
    return root;
}


function builtinSamplesFolder(context: vscode.ExtensionContext): string | undefined {
    // Packaged: <extension>/samples (copied by build-server.ps1); from the repository: VbsTools Studio's samples.
    return [path.join(context.extensionPath, "samples"), path.join(context.extensionPath, "..", "..", "src", "VbsTools.Studio", "samples")]
        .find(c => fs.existsSync(c));
}


function readSamplesFolder(dir: string): Node[] {
    let entries: fs.Dirent[];

    try {
        entries = fs.readdirSync(dir, { withFileTypes: true });
    } catch {
        return [];
    }

    const folders: Node[] = entries.filter(e => e.isDirectory()).map(e => ({
        kind: "folder", label: e.name, group: "samples", path: path.join(dir, e.name), children: readSamplesFolder(path.join(dir, e.name))
    } as FolderNode));

    const files: Node[] = entries.filter(e => e.isFile() && /\.vbs$/i.test(e.name)).map(e => {
        const file = path.join(dir, e.name);
        return { kind: "sample", label: e.name.replace(/\.vbs$/i, ""), text: () => readText(file) } as SampleNode;
    });

    const byLabel = (a: Node, b: Node) => a.label.localeCompare(b.label, undefined, { numeric: true });
    return [...folders.sort(byLabel), ...files.sort(byLabel)];
}


// ============================================================
// Helpers
// ============================================================

let builtinCache: Map<string, string> | undefined;

/** Shortcuts of the snippets that come with the extension (prefix → name). */
function builtinPrefixes(context: vscode.ExtensionContext): Map<string, string> {
    if (!builtinCache) {
        builtinCache = new Map();

        for (const s of readSnippets(path.join(context.extensionPath, "snippets", "vbscript.json"))) {
            builtinCache.set(s.prefix.toLowerCase(), s.name);
        }
    }

    return builtinCache;
}


function snippetUri(name: string): vscode.Uri {
    return vscode.Uri.from({ scheme: SNIPPET_SCHEME, path: "/" + encodeURIComponent(name) + ".vbs" });
}


function readText(file: string): string {
    try {
        return fs.readFileSync(file, "utf8").replace(/^﻿/, "");
    } catch {
        return "";
    }
}


function contentOf(node: Node): string | undefined {
    return node.kind === "sample" ? node.text() : node.kind === "snippet" ? snippetText(node.snippet.body) : undefined;
}


function firstLines(text: string, count: number): string {
    const lines = text.split(/\r?\n/);
    return lines.slice(0, count).join("\n") + (lines.length > count ? "\n…" : "");
}


function safeName(name: string): string {
    return name.replace(/[\\/:*?"<>|]/g, "_");
}


/** First comment line of the code, else "Snippet". */
function suggestName(text: string): string {
    const comment = text.split(/\r?\n/).map(l => l.trim()).find(l => l.startsWith("'"));
    return comment?.replace(/^'+\s*/, "").trim().substring(0, 60) || vscode.l10n.t("Snippet");
}


/** Initials of the name, made unique ("Read a Text File" → "rtf"). */
function suggestPrefix(name: string, store: SnippetStore): string {
    const words = name.split(/[^A-Za-z0-9]+/).filter(w => w && !/^(a|an|the|of|to|and|in|for)$/i.test(w));
    let prefix = (words.length > 1 ? words.map(w => w[0]).join("") : (words[0] ?? "snip").substring(0, 6)).toLowerCase();

    if (!/^[a-z_]/.test(prefix)) {
        prefix = "s" + prefix;
    }

    let candidate = prefix;

    for (let i = 2; store.all().some(s => s.prefix.toLowerCase() === candidate) || KEYWORDS.has(candidate); i++) {
        candidate = prefix + i;
    }

    return candidate;
}


function targetEditor(): vscode.TextEditor | undefined {
    const usable = (e: vscode.TextEditor) => e.document.uri.scheme !== PREVIEW_SCHEME && e.document.uri.scheme !== SNIPPET_SCHEME;
    const active = vscode.window.activeTextEditor;

    return active && usable(active)
        ? active
        : vscode.window.visibleTextEditors.find(e => e.document.languageId === "vbscript" && usable(e));
}


/**
 * Plain code as a snippet body: "$" is escaped (and a backslash only where it would
 * start an escape), so paths like C:\Windows stay readable when the snippet is edited.
 */
function escapeSnippet(text: string): string {
    return text.replace(/\r\n/g, "\n").replace(/\\(?=[$}\\])/g, "\\\\").replace(/\$/g, "\\$");
}


/** Snippet body as plain code: placeholders become their default text. */
export function snippetText(body: string): string {
    let text = body;
    let previous: string;

    do {
        previous = text;
        text = text
            .replace(/\$\{\d+\|([^,|]*)[^|]*\|\}/g, "$1")      // ${1|a,b|} → a
            .replace(/\$\{\d+:([^{}]*)\}/g, "$1")             // ${1:name} → name (innermost first)
            .replace(/\$\{\d+\}/g, "");
    } while (text !== previous);

    return text.replace(/(?<!\\)\$\d+/g, "").replace(/\\([$}\\])/g, "$1");
}


/** Snippets of a VS Code-format file (comments and trailing commas allowed; "folder" groups them). */
export function readSnippets(file: string): Snippet[] {
    let json: Record<string, { prefix?: string | string[]; body?: string | string[]; description?: string; folder?: string }>;

    try {
        json = JSON.parse(stripJsonComments(readText(file)));
    } catch {
        return [];
    }

    const result: Snippet[] = [];

    for (const [name, value] of Object.entries(json ?? {})) {
        if (!value || typeof value !== "object" || value.body === undefined) {
            continue;
        }

        const prefixes = Array.isArray(value.prefix) ? value.prefix : value.prefix ? [value.prefix] : [];
        result.push({
            name,
            prefix: prefixes[0] ?? name,
            body: Array.isArray(value.body) ? value.body.join("\n") : String(value.body),
            description: value.description ?? "",
            folder: value.folder ?? ""
        });
    }

    return result;
}


/** Removes line and block comments outside strings, and trailing commas (JSON with comments). */
export function stripJsonComments(text: string): string {
    let result = "";
    let inString = false;

    for (let i = 0; i < text.length; i++) {
        const c = text[i];

        if (inString) {
            result += c;

            if (c === "\\") {
                result += text[++i] ?? "";
            } else if (c === "\"") {
                inString = false;
            }
        } else if (c === "\"") {
            inString = true;
            result += c;
        } else if (c === "," && /^\s*[}\]]/.test(text.substring(i + 1, i + 200))) {
            // trailing comma: dropped
        } else if (c === "/" && text[i + 1] === "/") {
            while (i < text.length && text[i] !== "\n") { i++; }
            result += "\n";
        } else if (c === "/" && text[i + 1] === "*") {
            i += 2;
            while (i < text.length && !(text[i] === "*" && text[i + 1] === "/")) { i++; }
            i++;
        } else {
            result += c;
        }
    }

    return result;
}
