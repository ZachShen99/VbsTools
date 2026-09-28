# FAQ

## vbs files open as Visual Basic

VS Code's built-in Visual Basic support also claims `.vbs`. If the language mode in the status bar reads "Visual Basic", add to your settings (JSON):

```json
"files.associations": { "*.vbs": "vbscript" }
```

## No IntelliSense

1. Check that the language mode is **VBScript**.
2. Open the Output panel, choose **VBScript Language Server** and look for errors.
3. Run **VBScript: Restart Language Server**.
4. If an older version with the publisher `vbstools` is installed, uninstall it first (`code --uninstall-extension vbstools.vbstools-vbscript`); otherwise both run at once.

## No completion for a COM object

- The component may be registered for 32-bit only: switch to x86 and try again.
- The object comes from an application without a type library (such as WinCC): see [Type definitions and annotations](type-files.md).
- The variable's type cannot be inferred (a parameter, for instance): give it with an `' @type` comment.

## Garbled text

- **In the editor**: the file is ANSI in a code page other than the system's (a Japanese script on Chinese Windows, for instance). Click the encoding in the status bar → Reopen with Encoding, or set `vbscript.encoding.ansi`.
- **In cscript's output**: the file is UTF-8. Use the quick fix of the encoding warning to save it as UTF-16 LE or ANSI; see [Running scripts](running.md#file-encodings).
- **Thai, emoji and the like show as `?` in the output panel**: cscript converts its output to the console code page — a limit of cscript itself. The VbsTools host (`vbscript.run.host` = `vbstools`) shows them.

## "Input past end of file" while debugging

The script reads `WScript.StdIn`, and the Debug Console has no input. Without a `console` setting the extension uses a terminal; if `launch.json` says `"console": "internalConsole"`, change it to `"integratedTerminal"`.

## No values while debugging

- Stopped on an unhandled runtime error, the engine has already left the statement: only the call stack is available.
- Paused inside `WScript.Sleep`: step once to see the values on the next line.

## Differences from cscript

The VbsTools host was checked against cscript. Known differences:

- `.wsf` and `.vbe` files are not supported.
- `WScript.GetObject(file, progId)` with both arguments is not supported yet.
- `WScript.GetObject("winmgmts:")` works (cscript fails with "unknown protocol").
- Objects whose events are a pure vtable interface cannot be connected ("Could not connect object").
- The output panel and terminals show and accept any Unicode character (cscript is limited to the console code page).

`WScript.FullName` returning cscript.exe is intended: see [Running scripts](running.md#the-two-hosts).

## Antivirus warnings

The VbsTools host is a program that runs VBScript; an antivirus program may flag it by mistake. Add the extension folder (`%USERPROFILE%\.vscode\extensions\zach-shen.vbstools-vbscript-*`) to its exclusions.

## Reporting a problem

Please include:

- The extension version (shown in the Extensions view) and the Windows version
- A minimal script that shows the problem
- The contents of **VBScript Language Server** in the Output panel; for debugger problems, turn on `vbscript.debug.trace` and add the contents of **VBScript Debug Trace**
