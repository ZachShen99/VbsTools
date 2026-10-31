# Changelog

## 0.3.12

- Walkthrough: a new step "Debug scripts inside Altium DXP and other applications" takes a user from the setup
  (Process Debug Manager, Machine Debug Manager, Debugger Users group) through the `attach` configuration to the first
  breakpoint, with links to the environment check and to `launch.json`. English and Simplified Chinese pages; the step's
  title and description are translated into all six interface languages
- Languages: the interface (commands, settings, messages, diagnostics, debugger messages) is now also available in Traditional Chinese,
  Korean and German, besides English, Simplified Chinese and Japanese. The user guide stays in English and Simplified Chinese.
  The new translations have not been reviewed by native speakers
- Debugging: the setting `vbscript.debug.engine` now defaults to `auto`: the debugger uses the Process Debug Manager when pdm.dll is
  available (installed with Visual Studio or Office) and instruments the script otherwise. Before, the default was `instrumentation`;
  set it to get that back
- README: describes the debug engines and what the Process Debug Manager adds, debugging the scripts of other applications (attach, `takeOver`,
  the one-time setup, the check command, what to do when nothing happens), names assigned without `Dim` in completion, and the new settings

## 0.3.11

- Attach: the session now ends by itself when the script run it is attached to ends, for an application that registers each run of a script
  (like the Delphi test host). Before, the debug session stayed on, attached to a run that was over, until it was stopped by hand. An application
  that stays registered between scripts, like Altium DXP, keeps the session so that the next script is caught

## 0.3.10

- Attach: attaching to an application that already has a script loaded (the script is running, say stopped at a message box) now says in the
  Debug Console that the debugger does not stop by itself there: it stops at breakpoints reached later, or at Pause. Before, a script
  without such breakpoints just ran to its end and it was not clear why
- The messages of the attach (failures, skipped applications, files that were not loaded) are translated to Chinese and Japanese

## 0.3.9

- Attach: an application that the Machine Debug Manager still lists but that does not answer (typically the leftover process of an
  application that was closed — a `DXP.exe` without a window in the Task Manager) no longer blocks the attach for good. Each application
  has 5 seconds to answer; one that does not is skipped. When the application you asked for is not among the others, the message says how many
  were skipped and what to look for

## 0.3.8

- Attach: a failed attach no longer disconnects the debugger that was in the way. Before, the cleanup of the failed attempt disconnected
  whichever debugger was registered at the application, even though it was not ours
- Attach: new option `"takeOver": true` for the attach configuration. Altium DXP connects a debugger of its own when a script starts; to attach
  to a script that is already running (stopped at its first message box, say), the option disconnects that one and connects ours.
  Without it the attach is refused as before — attach first and run the script afterwards, or set the option

## 0.3.7

- Attach: when another debugger answers, the message lists the processes running now that may be it (a second script host, Visual Studio, …); with
  `VBSTOOLS_PDM_LOG` set the log also says whether the registered debugger is this session's own

## 0.3.6

- Attach: when the debugger cannot be connected to the application (0x80004005), the message now says what is registered there — another
  debugger that answers, one whose process has ended, or none (then the connection is retried a few times first). Before, it always
  claimed that "another debugger may be attached"

## 0.3.5

- Debugging: the Variables view now expands the objects a host application gives to its scripts (`MyApp`, `Board`, …) from the type files
  of the workspace (`*.vbstypes.json`, `vbscript.typeFiles`), not only from the built-in environments. Before, such objects stayed `{...}`
  unless the application was Altium

## 0.3.4

- IntelliSense: variables assigned without `Dim` (`x = 5`) now appear in name completion, local to their procedure or global at script level

## 0.3.3

- Attach: documented when the attach debugger can be used (`docs/DEBUGGER.md`, "When attaching works")
- New command **VBScript: Check Attach Debugging Environment**: checks `pdm.dll` (32/64-bit), the Machine Debug Manager service and
  the group *Debugger Users*, says what is missing, and can register an `mdm.exe` the user already has (only a Microsoft-signed
  one, as administrator) or add the user to the group. The extension still ships none of these components
- The same check runs when the extension starts and shows a notification when attach debugging cannot work yet; turn it off
  with the setting `vbscript.attach.checkOnStartup` or "Don't show again"

## 0.3.2

- Debugging (attach): when a breakpoint's file is not the one the application runs, the Debug Console says so — the breakpoint is
  bound to the same-named file in another folder, or the loaded script files are listed. A failure while binding breakpoints no
  longer leaves the application halted

## 0.3.1

- IntelliSense: built-in **Altium** environment (`' @environment Altium` or `"vbscript.environments": ["Altium"]`) for scripts that run
  inside Altium DXP: `PCBServer`, boards, iterators, primitives, `MkSet`, `ShowMessage`, `eTopLayer`, … — completion, hover,
  and no more "not defined" warnings for the host's own functions and constants
- Type definition files: new sections `functions` (global procedures / functions of the host) and `constants`, and `extends`
  for types that share members
- Debugging (PDM engine): a watch or hover of an undefined name or a missing member is now an error, not a value
- Debugging: objects a host injects without type information (Altium's `Board`) are expandable in the Variables view — the
  members come from the script's environment type, each value is read from the running application; `[Methods]` lists the
  methods. Attach configuration: new `environments` option (default `Altium` for `DXP.exe`)

- Debugging: **attach to another application** — `"request": "attach"` debugs the scripts that run inside an application that
  registers them with the Process Debug Manager (Altium DXP). Needs the Machine Debug Manager (`mdm.exe /regserver`),
  membership of the Debugger Users group, and the application started afterwards. See `docs/DEBUGGER.md`.
- Debugging: new engine setting `vbscript.debug.engine` (`instrumentation` default, `pdm`, `auto`). The Process
  Debug Manager engine (pdm.dll, installed with Visual Studio or Office) stops at runtime errors with the
  variables intact, pauses loops without statements and shows the engine's own call stack; it falls back to
  instrumentation when pdm.dll is not available. See `docs/DEBUGGER.md`.
- Debugging (PDM engine): arrays, objects and Dictionaries are summarized (`Variant(0 To 2)`, a class name, `Dictionary (Count=2)`),
  a Dictionary lists its `key = item` entries, and the methods of an object are listed under `[Methods]`
- Debugging (PDM engine): `Execute` / `ExecuteGlobal` code can be stepped into and read (virtual source); stack frames carry the
  statement's column range; watches and hovers are evaluated without side effects
- Debugging (PDM engine): at a runtime error, `!retry` and `!skip` in the Debug Console run the statement again / move on
  past it; "Jump to Cursor" (Set Next Statement) moves the execution point

## 0.2.1

- Debugging: variables inside `With CreateObject(...)` / `With New Cls` blocks — the object shows as a **With**
  scope in the Variables view, and `.Member` expressions work in Watch, Debug Console, and conditional
  breakpoints / logpoints inside the block (the reference is released at `End With`, so `Class_Terminate`
  still runs on time)
- Debugging: a `Stop` statement breaks into the debugger (known limit: `If … Then Stop` as the very last
  statement of a procedure stops in the caller)
- Debugging: the objects of unfinished `With` blocks are released when the script ends or fails, so
  `Class_Terminate` runs as it does under cscript
- IntelliSense: .NET components registered with `regasm /tlb` (no `TypeLib` key under the CLSID, server
  `mscoree.dll`) now get completion — the type library next to the assembly (`CodeBase`) is found
- Hover: removed the redundant "Member of …" line
- File icon for `.vbs` / `.vbe` files
- Smaller package: no debug symbols for the language server, compressed single file

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
