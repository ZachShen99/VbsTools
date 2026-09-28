# Changelog

## 0.2.0

- Walkthrough **Get Started with VbsTools** (Help → Welcome): a sample script, running, debugging, the
  Code Library, host environments, the user guide — in English, Simplified Chinese and Japanese
- Screenshots and animations in the README (the extension's page), the user guide and the walkthrough
- Run Script and debugging (F5) work while another page has the focus (the walkthrough, a Markdown preview):
  they use the VBScript editor that is visible

## 0.1.1

- File encodings: ANSI scripts (GBK, Shift JIS, …) that VS Code shows garbled are reopened with the system
  code page; a warning when the encoding makes cscript / wscript read the script wrong (UTF-8 with text
  that is not ASCII, UTF-8 with BOM), with quick fixes *Save as UTF-16 LE* / *Save as ANSI*; command
  **VBScript: Save with Encoding…**; settings `vbscript.encoding.ansi`, `vbscript.encoding.check`
- *Run Without Debugging* (the Run button, the Run menu) runs like **VBScript: Run Script**: with
  `vbscript.run.host` (cscript was ignored), output in the *VBScript* panel, errors in Problems;
  *Run Script* and *Run Script with Arguments…* in the Run button's menu
- User interface in Simplified Chinese and Japanese — follows VS Code's display language: commands,
  settings, Code Library, messages, diagnostics and quick fixes, documentation of built-in functions and
  WScript, debugger messages, type definition file errors
- Debugging (F5) a script that reads `WScript.StdIn` runs it in a terminal (the Debug Console has no input:
  "Input past end of file"), following `vbscript.run.console`
- Completion for `WScript.StdIn` / `StdOut` / `StdErr` (ReadLine, WriteLine, AtEndOfStream, …)
- VbsTools Script Host in a terminal (Run Script with StdIn, debugging): any character is shown and can be
  typed (Thai, emoji, …), not only those of the console code page
- Output of `WScript.StdOut` and `WScript.StdErr` keeps its order (Run Script with the VbsTools Script Host,
  debugging in the Debug Console, where stderr shows in red)
- User guide in English and Simplified Chinese: **VBScript: Open Documentation** (also in the Code Library's
  **…** menu)
- README (the extension's page) in English and Simplified Chinese
- Requires VS Code 1.100 or later

## 0.1.0 (pre-release)

First release.

- IntelliSense for COM objects: members, chains, `For Each`, `With`, default members, script classes,
  ProgID completion, parameter info, hover, parameter name hints
- Type definition files (`*.vbstypes.json`) for applications that run VBScript; built-in WinCC V7/V8
- Type annotations (`' @type`, `' @param`, `' @returns`) and included scripts (`Include "…"`, `' @include`)
- Diagnostics: syntax errors (vbscript.dll), `Option Explicit`, undefined procedures, argument counts,
  unknown classes, unused variables; quick fixes
- Outline, definition, references, rename, call hierarchy, reference CodeLens, semantic highlighting,
  formatting, folding (including `'#Region`)
- Run with the bundled cscript-compatible VbsTools Script Host (x64 / x86) or cscript; arguments per script
- Debugger: breakpoints (conditions, hit counts, logpoints), stepping, pause, variables (editable), watch,
  Debug Console with completion, runtime and handled errors, COM events
- Code Library: VbsEdit / VbsTools Studio samples and My Snippets (shared with VbsTools Studio)
