# 调试

打开脚本按 **F5** 即可开始调试，不需要 `launch.json`。调试时脚本在 VbsTools 脚本宿主中运行，平台（x64 / x86）和参数与"运行脚本"相同。

## 断点

| 类型 | 设置方法 |
|---|---|
| 普通断点 | 单击行号左侧，或 F9 |
| 条件断点 | 右键断点 → 编辑断点 → 表达式，例如 `i > 10 And name = "abc"` |
| 命中次数 | 编辑断点 → 命中次数：`5`（第 5 次）、`>= 5`、`> 5`、`< 5`、`<= 5`、`% 5`（每 5 次） |
| 日志点 | 编辑断点 → 日志消息，例如 `i = {i}, name = {name}`；不暂停，把消息输出到调试控制台 |

没有可执行代码的行（`Dim`、`End If`、空行等）上的断点会移到下一条可执行语句。

## 单步执行

| 操作 | 快捷键 |
|---|---|
| 继续 | F5 |
| 单步跳过 | F10 |
| 单步进入 | F11 |
| 单步跳出 | Shift+F11 |
| 暂停 | F6（在 `WScript.Sleep` 中等待时也能暂停） |
| 停止 | Shift+F5 |

## 查看和修改值

- **变量**视图：局部变量和全局变量。字符串带引号显示；`Empty` / `Null` / `Nothing` 原样显示；数组显示为 `Variant(0 To 3)` 并可展开；COM 对象按类型信息列出属性；脚本类的对象列出公共字段和 `Property Get`（不会调用方法）；Dictionary 显示为 键 → 值。
- **修改值**：双击变量，输入新值或表达式（在当前作用域中计算，对象会自动用 `Set` 赋值）。
- **监视**：添加任意表达式，每次暂停时重新计算。
- **悬停**：暂停时鼠标停在变量上显示它的值。
- `Err`、`Err.Number`、`Err.Description` 在监视和条件中看到的是脚本自己的错误状态。

## 调试控制台

暂停时在调试控制台中输入，规则与 VB 的"立即窗口"相同：

| 输入 | 效果 |
|---|---|
| `? x * 2` | 计算表达式并显示结果 |
| `x = 5` | 执行赋值 |
| `Call Foo(1)` | 执行语句 |
| 其他 | 当作表达式计算 |

输入时有自动补全，和编辑器中的补全一样能理解 COM 对象。

脚本的输出（`WScript.Echo`、`StdOut`、`StdErr`）也显示在调试控制台中，`StdErr` 为红色。脚本使用 `WScript.StdIn` 时，会自动改在终端中运行，这样才能输入；断点和单步照常可用。

## 异常

"断点"视图中有两个选项：

| 选项 | 默认 | 说明 |
|---|---|---|
| 运行时错误 | 开 | 发生未处理的运行时错误时，在出错的行停下 |
| 已处理的错误（On Error Resume Next） | 关 | `On Error Resume Next` 生效期间发生错误时，在出错的行停下 |

在未处理的错误处停下时，脚本引擎已经离开了出错的语句，所以无法查看变量值，只能看到调用堆栈。

## COM 事件

`WScript.CreateObject("ProgID", "前缀_")` 和 `WScript.ConnectObject` 连接的事件在调试时正常触发，也可以在事件处理过程中设断点。暂停期间到达的事件会排队，继续运行后再依次处理。

## launch.json

需要固定参数或其他选项时，可以在 `.vscode/launch.json` 中添加配置（在该文件中输入 `VBScript` 会有代码片段）：

```jsonc
{
  "version": "0.2.0",
  "configurations": [
    {
      "type": "vbscript",
      "request": "launch",
      "name": "Debug VBScript",
      "program": "${file}",
      "args": ["in.txt", "/verbose"],     // WScript.Arguments
      "cwd": "${fileDirname}",            // 工作目录
      "platform": "x86",                  // x64 / x86，默认同 vbscript.run.platform
      "stopOnEntry": false,               // 在第一条语句处停下
      "console": "integratedTerminal"     // internalConsole / integratedTerminal / externalTerminal
    }
  ]
}
```

不写 `console` 时按设置 `vbscript.run.console` 决定：脚本用到 `WScript.StdIn` 就用终端，否则用调试控制台。

## 限制

- 单步以行为单位；一行里有多条语句（`a = 1: b = 2`）时一次执行完。
- 没有语句的循环（`Do : Loop`）无法暂停。
- 不支持 `.wsf` 和 `.vbe` 文件。
