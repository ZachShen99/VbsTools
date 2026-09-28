# VbsTools documentation

VbsTools turns VS Code into a complete VBScript (`.vbs`) environment: IntelliSense that understands COM objects, diagnostics, a debugger, a cscript-compatible script host and a code library.

> Windows only. Type information comes from the COM type libraries registered on your machine; scripts run in the script host bundled with the extension.

## Contents

| Chapter | Topics |
|---|---|
| [Getting started](getting-started.md) | Installing, a first script, running and debugging |
| [Editing and IntelliSense](editing.md) | COM completion, hover, parameter info, diagnostics, quick fixes, navigation, formatting |
| [Running scripts](running.md) | Run commands, arguments, output panel and terminal, the two hosts, 32/64-bit, file encodings |
| [Debugging](debugging.md) | Breakpoints, stepping, variables, watch, Debug Console, exceptions, launch.json |
| [Type definitions and annotations](type-files.md) | Host environments such as WinCC, `*.vbstypes.json`, `@type` comments, included scripts |
| [Code Library](code-library.md) | Samples, My Snippets, shortcuts |
| [Settings](settings.md) | All settings |
| [FAQ](faq.md) | Troubleshooting, differences from cscript |

## Keyboard shortcuts

| Action | Shortcut |
|---|---|
| Run script | Ctrl+F5 |
| Run with arguments | Ctrl+Shift+F5 |
| Stop | Shift+F5 |
| Debug | F5 |
| Toggle breakpoint | F9 |
| Step over / into / out | F10 / F11 / Shift+F11 |
| Quick fix | Ctrl+. |
| Go to definition / Find all references / Rename | F12 / Shift+F12 / F2 |
| Format document | Shift+Alt+F |

Come back here any time with **VBScript: Open Documentation** in the Command Palette (Ctrl+Shift+P).
