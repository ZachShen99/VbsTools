# 类型描述与注释

COM 类型库之外，还有三种方式告诉插件对象的类型：宿主环境（类型描述文件）、类型注释、Include 文件。

## 宿主环境：WinCC

有些 VBScript 在软件内部运行，软件会注入全局对象（例如 WinCC 的 `HMIRuntime`）。这些对象没有注册的类型库，需要用类型描述文件（`*.vbstypes.json`）来描述。插件内置了 **WinCC V7 / V8（经典版）**。

启用方式二选一：

- 所有脚本都启用：设置 `"vbscript.environments": ["WinCC"]`
- 只对单个脚本启用：在脚本中写注释 `' @environment WinCC`

启用后：

```vbscript
' @environment WinCC
Sub OnClick(ByVal Item)
    Item.                                   ' 画面对象的属性
    HMIRuntime.Tags("Motor1").Read          ' 变量对象
    HMIRuntime.Screens("Base").ScreenItems("Button1").Visible = False
End Sub
```

内置的 WinCC 描述是根据 WinCC VBS 参考整理的，个别成员可能与你的版本不同。发现缺漏时，可以在工作区里另写一个类型描述文件补充。

## 自己编写类型描述文件

工作区中任意位置的 `*.vbstypes.json` 会自动生效，保存后立即重新加载。放在工作区外的文件用设置 `vbscript.typeFiles` 指定（文件或文件夹，可用 `%USERPROFILE%` 等环境变量）。编辑时有 JSON Schema 校验和补全，错误显示在"问题"面板。

```jsonc
{
  "name": "Demo",                                  // 必填：环境名
  "description": "Demo application",
  "globals":    { "App": "Application" },          // 全局对象：名称 → 类型
  "parameters": { "Item": "Control" },             // 按参数名指定类型（事件过程）
  "progIds":    { "Demo.Counter": "Counter" },     // CreateObject("Demo.Counter") 返回的类型
  "enums":      { "Mode": { "Fast": 1, "Safe": 2 } },
  "types": {
    "Application": {
      "description": "根对象",
      "members": {
        "Version": "String",                        // 字符串：该类型的属性
        "Items":   { "type": "Items", "readonly": true, "doc": "所有项目" },
        "Log":     { "kind": "method", "params": [ { "name": "Text", "type": "String" } ] },
        "Run":     { "kind": "method", "type": "Long",               // 有 type 的方法 = Function
                     "params": [ { "name": "Name", "type": "String" },
                                 { "name": "Mode", "type": "Mode", "default": 1 } ] }
      }
    },
    "Items": {
      "element": "Item",                            // For Each 的元素类型
      "members": {
        "Item":  { "type": "Item", "default": true, "params": [ { "name": "Index", "type": "Variant" } ] },
        "Count": "Long"
      }
    },
    "Item":    { "open": true, "members": { "Name": "String" } },     // open：还有未列出的成员
    "Counter": { "members": { "Value": "Long" } },
    "Control": { "members": { "Visible": "Boolean" } }
  }
}
```

- **类型名**：VBScript 类型（`Variant String Long Integer Boolean Double Single Date Currency Byte Object`）、本文件中的类型或枚举，或 `progid:Excel.Application`（引用已注册的 COM 类型）。
- **成员**：`kind` 为 `property`（默认）或 `method`；`default: true` 表示默认成员（`obj("x")` 就是调用它）；参数可以写 `optional`、`default`（有默认值即为可选）、`paramArray`。
- **枚举**：用作参数或属性的类型时，补全会列出这些值；插入的是数值，因为 VBScript 运行时不认识这些名字。
- 文件中允许注释和末尾的逗号。

## 类型注释

插件推断不出类型时（例如对象作为参数传入），可以用注释说明：

```vbscript
' @type f As Scripting.File
For Each f In files
    f.                          ' File 的成员
Next

' @param ws As Excel.Worksheet
' @returns Scripting.Folder
Function Export(ws)
    ws.                         ' Worksheet 的成员
    Set Export = CreateObject("Scripting.FileSystemObject").GetFolder("C:\")
End Function

Export(sheet).                  ' Folder 的成员
```

| 注释 | 写在哪里 | 作用 |
|---|---|---|
| `' @type 变量 As 类型` | 变量所在的作用域内任意位置 | 变量的类型 |
| `' @param 参数 As 类型` | 紧挨在 Sub / Function 上方的注释块 | 参数的类型 |
| `' @returns 类型` | 紧挨在 Function / Property Get 上方的注释块 | 返回值的类型 |

`As` 可以省略。类型可以是 ProgID（`Scripting.FileSystemObject`）、已注册类型库的 `库名.类型`（`Excel.Worksheet`、`Scripting.File`）、脚本中定义的类，或当前宿主环境中的类型。写错的类型名会显示警告。

## Include 文件

VBScript 没有 include 语句，通常用 `ExecuteGlobal` 读入另一个文件：

```vbscript
Include "lib\common.vbs"

Sub Include(path)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    ExecuteGlobal fso.OpenTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), path)).ReadAll
End Sub
```

插件能识别以下写法，并读取被包含的文件：被包含文件中的过程、类和全局变量会出现在补全中，不会被报告为未定义；F12 和 Ctrl+单击可以打开它们。

- 调用脚本中带参数、内部使用了 `ExecuteGlobal` 的过程，参数是字符串常量
- 调用名为 `Include`、`IncludeFile`、`IncludeScript`、`Import` 的过程，参数是字符串常量
- 注释 `' @include lib\common.vbs`（路径也可以加引号）

相对路径以**当前脚本所在的文件夹**为起点。找不到文件时会显示警告。被包含的文件中再包含的文件也会读取（最多 4 层）。
