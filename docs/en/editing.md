# Editing and IntelliSense

## IntelliSense for COM objects

The extension reads members from the COM type libraries registered on your machine and infers the types of variables from the code:

```vbscript
Set fso = CreateObject("Scripting.FileSystemObject")
fso.                                ' members of FileSystemObject
fso.GetFolder("C:\").Files.         ' chains: members of the Files collection
For Each f In fso.GetFolder("C:\").Files
    f.                              ' element type of For Each: File
Next
With fso.GetFile("a.txt")
    .                               ' With blocks
End With
```

Types come from `CreateObject` / `GetObject` / `WScript.CreateObject`, `Set x = …` assignments, function return values, `For Each` element types, `With` blocks, default members (`dict("key").`) and classes of the script (`Set p = New Person`).

| Feature | Description |
|---|---|
| ProgID completion | Inside `CreateObject("`: the ProgIDs registered on your machine |
| Parameter info | On `(` and `,`: the parameter list, current parameter highlighted |
| Hover | Signature and description of members and variables |
| Parameter name hints | Names at calls, e.g. `fso.OpenTextFile("a.txt", IOMode: 8, Create: True)`; see `vbscript.inlayHints.parameterNames` |
| Enum values | For a parameter of an enum type, completion lists its values and inserts the number (VBScript does not know type library constants at run time) |
| Built-ins | VBScript functions such as `MsgBox` and `InStr`, and `WScript`, `WScript.StdIn` / `StdOut`, `Err` |

Where the type cannot be inferred (an object passed as a parameter, for instance), give it in a comment: see [type annotations](type-files.md#type-annotations).

## Snippets

Type a prefix and choose it from the suggestions (or press Tab):

`fso` `dict` `shell` `net` `readfile` `writefile` `xl` `word` `fe` `for` `if` `ife` `sel` `do` `func` `sub` `cls` `wmi` `oe` `http` `ado` `args`

Your own snippets with their own shortcuts: see [Code Library](code-library.md).

## Diagnostics

Checked as you type; the script is never run:

| Check | Severity |
|---|---|
| Syntax errors (checked by vbscript.dll, messages in the Windows language) | Error |
| Undeclared variables under `Option Explicit` | Warning |
| Calls to a Sub / Function that does not exist | Warning |
| Wrong number of arguments | Warning |
| Unknown class after `New` | Warning |
| Unused variables (shown faded) | Hint |
| Included script not found, unknown type in a type annotation | Warning |
| An encoding that makes cscript read the file wrong (see [Running scripts](running.md#file-encodings)) | Warning |

### Quick fixes (Ctrl+.)

- Declare an undeclared variable with `Dim`, or all of them at once
- Remove an unused variable
- Save the file as UTF-16 LE / ANSI (encoding warning)

## Navigation

| Feature | How |
|---|---|
| Outline | Outline view in the Explorer, or Ctrl+Shift+O |
| Go to definition | F12 or Ctrl+click; also into included scripts |
| Find all references | Shift+F12 |
| Rename | F2; members of script classes are renamed at `obj.Member` too |
| Call hierarchy | Shift+Alt+H |
| Reference counts | *N references* above procedures and classes, click to list them; turn off with `vbscript.codeLens.references` |
| Semantic highlighting | Variables, parameters, procedures, classes, COM members and host objects in distinct colors |

## Formatting and folding

- **Format Document** (Shift+Alt+F) fixes the indentation. By default it also writes keywords and built-ins in their usual case (`End If`, `MsgBox`, `vbCrLf`) and every name of the script like its declaration. For indentation only, set `vbscript.format.keywordCase` to `preserve`.
- **Folding** by blocks (`If`, `For`, `Sub`, `Class`, …) and by region comments:

  ```vbscript
  '#Region Initialization
  ...
  '#End Region
  ```

  Region names also show in the minimap.
