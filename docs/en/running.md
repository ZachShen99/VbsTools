# Running scripts

## Run commands

| Command | Shortcut | Description |
|---|---|---|
| VBScript: Run Script | Ctrl+F5 | Runs the current script (unsaved changes are saved first) |
| VBScript: Run Script with Arguments… | Ctrl+Shift+F5 | Asks for the command line, then runs |
| VBScript: Stop Script | Shift+F5 | Ends the running script |

The Run button (▷) at the top right of the editor and *Run Without Debugging* in the Run menu do the same; the button's menu also has *Run Script* and *Run Script with Arguments*.

### Arguments

The command line entered for *Run with Arguments* is remembered per script and used by later runs (Ctrl+F5) and by debugging (F5). Quote arguments that contain spaces:

```text
in.txt /out:"result file.txt" /verbose
```

The script reads them from `WScript.Arguments`; `/name:value` arguments also from `WScript.Arguments.Named`.

### Where the output goes

| Case | Output |
|---|---|
| Most scripts | The *VBScript* output panel; runtime and syntax errors in the Problems panel, the cursor on the error |
| The script uses `WScript.StdIn` | A VS Code terminal (the output panel has no input) |

`vbscript.run.console` can fix it to `output` or `terminal`.

## The two hosts

| | VbsTools Script Host (default) | cscript |
|---|---|---|
| `vbscript.run.host` | `vbstools` | `cscript` |
| Behavior | Like cscript (checked against it) | Part of Windows |
| Characters in the output panel | Any (UTF-8 output) | Characters outside the console code page show as `?` |
| Typing in a terminal | Any character (Thai, emoji, …) | Only characters of the console code page |
| Order of StdOut and StdErr | As written | May be mixed up in the output panel |
| Debugging | Yes (debugging always uses it) | — |

Like cscript, `WScript.FullName` and `WScript.Path` return the real cscript.exe (many scripts check it to find out whether they run in a console), and the exit code is 0 after a runtime error, 1 after a syntax error and n after `WScript.Quit n`. Known differences: see the [FAQ](faq.md#differences-from-cscript).

## 32-bit / 64-bit

Switch with **VBScript x64 / x86** in the status bar or **VBScript: Switch 64-bit / 32-bit**; the setting is `vbscript.run.platform`. COM components registered for 32-bit only need x86: cscript from SysWOW64, or the x86 VbsTools host.

## File encodings

cscript and wscript read only two encodings: **UTF-16 LE with BOM**, or the **system ANSI code page** (GBK on Simplified Chinese Windows, Shift JIS on Japanese Windows, Windows-1252 on Western Windows). They do not understand UTF-8 — which is what VS Code saves by default. So the extension:

1. **Recognizes ANSI files**: scripts saved as ANSI by Notepad, VbsEdit and others, which VS Code would show garbled as UTF-8, are reopened with the system code page. Another code page: `vbscript.encoding.ansi`.
2. **Warns** when a file is UTF-8 with non-ASCII text (accented letters, Chinese, …), or UTF-8 with BOM:
   - UTF-8: the VbsTools host reads it correctly, but cscript, double-clicking and scheduled tasks show garbled text;
   - UTF-8 with BOM: cscript cannot run it at all ("Invalid character").
3. **Converts**: Ctrl+. on the warning offers *Save as UTF-16 LE* (runs anywhere) and *Save as ANSI*; **VBScript: Save with Encoding…** does the same at any time.

| Encoding | cscript / wscript | VbsTools host | Notes |
|---|---|---|---|
| UTF-16 LE (BOM) | ✓ | ✓ | Recommended; git treats it as binary |
| ANSI (Windows-1252, GBK, …) | ✓ (systems with the same code page) | ✓ | Garbled on systems with another language |
| UTF-8 | Non-ASCII text garbled | ✓ | |
| UTF-8 with BOM | ✗ does not run | ✓ | |

To turn the warnings off, set `vbscript.encoding.check` to `false`.

## The VbsTools host on the command line

The host is in the extension folder under `server\host\x64\` and `server\host\x86\`; it is used like cscript:

```text
VbsTools.ScriptHost.exe [//options] script.vbs [arguments]
```

| Option | Description |
|---|---|
| `//B` | Batch mode (`WScript.Interactive = False`) |
| `//T:nn` | Stop the script after nn seconds |
| `//W` | Show `WScript.Echo` and errors as message boxes, like wscript |
| `//UTF8` | Standard input and output in UTF-8 (for editors) |
| `//ErrToOut` | Standard error goes to standard output (keeps their order) |
| `//Nologo` | Accepted for compatibility; no logo is shown |
