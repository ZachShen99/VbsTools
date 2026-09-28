# 常见问题

## vbs 文件显示为 Visual Basic

VS Code 自带的 Visual Basic 支持也会认领 `.vbs`。右下角语言模式显示 "Visual Basic" 时，在设置（JSON）中加入：

```json
"files.associations": { "*.vbs": "vbscript" }
```

## 没有智能提示

1. 确认右下角语言模式是 **VBScript**。
2. 打开"输出"面板，下拉选择 **VBScript Language Server**，查看是否有错误。
3. 执行 **VBScript: 重启语言服务器**。
4. 如果装过旧版本的插件（发布者为 `vbstools`），先卸载它：`code --uninstall-extension vbstools.vbstools-vbscript`，否则两个插件会同时运行。

## 某个 COM 对象没有补全

- 该组件可能只注册了 32 位：把平台切换为 x86 后再试。
- 对象来自没有类型库的软件（例如 WinCC）：见[类型描述与注释](type-files.md)。
- 变量的类型推断不出来（例如作为参数传入）：用 `' @type` 注释说明。

## 中文显示为乱码

- **编辑器中乱码**：文件是 ANSI 编码，但不是系统代码页（例如在中文系统上打开日文脚本）。点右下角的编码 → 通过编码重新打开，或设置 `vbscript.encoding.ansi`。
- **cscript 运行结果乱码**：文件是 UTF-8。按编码警告的快速修复另存为 UTF-16 LE 或 ANSI，见[运行脚本](running.md#文件编码)。
- **输出面板中泰文、emoji 显示为 `?`**：cscript 输出时会转换成控制台代码页，这是 cscript 本身的限制。改用 VbsTools 宿主（`vbscript.run.host` = `vbstools`）即可。

## 调试时提示"输入超出了文件尾"

脚本读取 `WScript.StdIn`，但调试控制台无法输入。不写 `console` 时插件会自动改用终端；如果 `launch.json` 里写了 `"console": "internalConsole"`，改为 `"integratedTerminal"`。

## 调试时无法查看变量

- 在未处理的运行时错误处停下时，脚本引擎已经离开出错的语句，只能看调用堆栈。
- 在 `WScript.Sleep` 中暂停时也看不到变量，单步一次后在下一行即可查看。

## 与 cscript 的差异

VbsTools 宿主与 cscript 逐项对照过，已知差异：

- 不支持 `.wsf` 和 `.vbe` 文件。
- `WScript.GetObject(文件, ProgID)` 同时给出两个参数的写法暂不支持。
- `WScript.GetObject("winmgmts:")` 可以使用（cscript 会报"未知协议"）。
- 只有纯 vtable 事件接口的对象无法连接事件（报"无法连接对象"）。
- 输出面板和终端中能显示和输入任意 Unicode 字符（cscript 受控制台代码页限制）。

`WScript.FullName` 返回 cscript.exe 的路径是故意的，见[运行脚本](running.md#两种宿主)。

## 杀毒软件拦截

VbsTools 宿主是一个运行 VBScript 的程序，个别杀毒软件可能会误报。把插件目录（`%USERPROFILE%\.vscode\extensions\zach-shen.vbstools-vbscript-*`）加入白名单即可。

## 报告问题

请附上以下信息：

- 插件版本（扩展视图中可以看到）、Windows 版本
- 能重现问题的最小脚本
- "输出"面板中 **VBScript Language Server** 的内容；调试相关的问题，打开 `vbscript.debug.trace` 后附上 **VBScript Debug Trace** 的内容
