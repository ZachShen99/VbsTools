# Settings

Search for `vbscript` in the Settings editor to see all of them.

## Editing

| Setting | Default | Description |
|---|---|---|
| `vbscript.format.keywordCase` | `pascal` | Whether formatting also normalizes the case of keywords and names; `preserve`: indentation only |
| `vbscript.inlayHints.parameterNames` | `literals` | Parameter name hints: `none`; `literals` (before literal arguments only); `all` |
| `vbscript.codeLens.references` | `true` | Reference counts above procedures and classes |
| `vbscript.environments` | `[]` | Host environments for every script, e.g. `["WinCC"]` |
| `vbscript.typeFiles` | `[]` | Type definition files or folders outside the workspace |

## Running and debugging

| Setting | Default | Description |
|---|---|---|
| `vbscript.run.host` | `vbstools` | The host that runs scripts: `vbstools` or `cscript`; debugging always uses the VbsTools host |
| `vbscript.run.platform` | `x64` | `x64` or `x86` (COM components registered for 32-bit only) |
| `vbscript.run.console` | `auto` | Where output goes: `auto` (a terminal when the script uses StdIn), `output`, `terminal`; also the default for debugging |
| `vbscript.run.scriptHostPath` | empty | Folder of your own VbsTools host build (with `x64\` and `x86\`) |

## File encodings

| Setting | Default | Description |
|---|---|---|
| `vbscript.encoding.ansi` | `auto` | Code page of ANSI scripts; `auto`: the system ANSI code page |
| `vbscript.encoding.check` | `true` | Warn when the encoding makes cscript read the script wrong |

## Code Library

| Setting | Default | Description |
|---|---|---|
| `vbscript.library.samplesFile` | empty | VbsEdit's samples file; empty: looked for in the VbsEdit installation folder |

## Troubleshooting

| Setting | Default | Description |
|---|---|---|
| `vbscript.languageServer.path` | empty | Path of your own language server build |
| `vbscript.languageServer.trace` | `false` | Log every completion request in the *VBScript Language Server* output channel (restart the server after changing it) |
| `vbscript.debug.trace` | `false` | Log all debugger messages in the *VBScript Debug Trace* output channel |

## Related settings

- `"files.associations": { "*.vbs": "vbscript" }` — when `.vbs` files open as Visual Basic
- `editor.inlayHints.enabled` — VS Code's switch, also for parameter name hints
