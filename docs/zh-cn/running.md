# 运行脚本

## 运行命令

| 命令 | 快捷键 | 说明 |
|---|---|---|
| VBScript: 运行脚本 | Ctrl+F5 | 运行当前脚本（未保存的修改会先保存） |
| VBScript: 带参数运行脚本… | Ctrl+Shift+F5 | 先输入命令行参数再运行 |
| VBScript: 停止脚本 | Shift+F5 | 结束正在运行的脚本 |

编辑器右上角的运行按钮（▷）和"运行"菜单中的"以非调试模式运行"效果相同；按钮的下拉菜单里也有"运行脚本"和"带参数运行脚本"。

### 参数

"带参数运行"输入的命令行会按脚本分别记住，之后的 Ctrl+F5 和 F5 调试都会使用。含空格的参数加引号：

```text
in.txt /out:"result file.txt" /verbose
```

脚本中通过 `WScript.Arguments` 读取，`/name:value` 形式的参数也可以从 `WScript.Arguments.Named` 读取。

### 输出位置

| 情况 | 输出到 |
|---|---|
| 一般脚本 | "VBScript" 输出面板；运行时错误和语法错误显示在"问题"面板，光标跳到出错的行 |
| 脚本使用了 `WScript.StdIn` | VS Code 终端（输出面板无法输入） |

可以用设置 `vbscript.run.console` 固定为 `output` 或 `terminal`。

## 两种宿主

| | VbsTools 脚本宿主（默认） | cscript |
|---|---|---|
| 设置 `vbscript.run.host` | `vbstools` | `cscript` |
| 行为 | 与 cscript 一致（逐项对照验证） | Windows 自带 |
| 输出面板中的字符 | 任意字符（UTF-8 输出） | 控制台代码页中没有的字符显示为 `?` |
| 终端中输入 | 任意字符（泰文、emoji 等） | 只能输入控制台代码页中的字符 |
| StdOut 与 StdErr 的顺序 | 保持写入顺序 | 输出面板中可能错乱 |
| 调试 | 支持（调试始终使用它） | — |

与 cscript 保持一致的地方：`WScript.FullName` 和 `WScript.Path` 返回真正的 cscript.exe 路径（很多脚本靠它判断自己是否在控制台中运行）；运行时错误后退出码为 0，语法错误为 1，`WScript.Quit n` 为 n。已知差异见[常见问题](faq.md#与-cscript-的差异)。

## 32 位 / 64 位

状态栏的 **VBScript x64 / x86** 或命令 **VBScript: 切换 64 位 / 32 位** 用来切换，对应设置 `vbscript.run.platform`。只注册了 32 位的 COM 组件需要 x86：使用 SysWOW64 下的 cscript，或 x86 版 VbsTools 宿主。

## 文件编码

cscript 和 wscript 只能读两种编码：**带 BOM 的 UTF-16 LE**，或**系统 ANSI 代码页**（简体中文 Windows 为 GBK，日文 Windows 为 Shift JIS）。它们不认识 UTF-8，而 VS Code 默认把文件保存为 UTF-8。插件对此做了三件事：

1. **自动识别 ANSI 文件**：用记事本、VbsEdit 等保存的 ANSI 脚本，VS Code 会误当成 UTF-8 显示成乱码，插件会用系统代码页重新打开。需要其他代码页时设置 `vbscript.encoding.ansi`。
2. **编码警告**：文件为 UTF-8 且含有非 ASCII 字符（中文、日文等），或为带 BOM 的 UTF-8 时，显示警告：
   - UTF-8：VbsTools 宿主能正确读取，但 cscript、双击运行、计划任务会显示乱码；
   - 带 BOM 的 UTF-8：cscript 直接报"无效字符"。
3. **转换编码**：警告处按 Ctrl+. 选择"另存为 UTF-16 LE"（任何系统都能运行）或"另存为 ANSI"；也可以随时执行 **VBScript: 以指定编码保存…**。

| 编码 | cscript / wscript | VbsTools 宿主 | 说明 |
|---|---|---|---|
| UTF-16 LE（带 BOM） | ✓ | ✓ | 推荐；缺点是 git 会把它当作二进制文件 |
| ANSI（GBK 等） | ✓（同一代码页的系统） | ✓ | 换到其他语言的系统上会乱码 |
| UTF-8 | 非 ASCII 字符乱码 | ✓ | |
| UTF-8（带 BOM） | ✗ 无法运行 | ✓ | |

不需要这些警告时，设置 `vbscript.encoding.check` 为 `false`。

## 在命令行使用 VbsTools 宿主

宿主程序位于插件目录的 `server\host\x64\` 和 `server\host\x86\`，用法与 cscript 相同：

```text
VbsTools.ScriptHost.exe [//选项] script.vbs [参数]
```

| 选项 | 说明 |
|---|---|
| `//B` | 批处理模式（`WScript.Interactive = False`） |
| `//T:nn` | nn 秒后停止脚本 |
| `//W` | 像 wscript 一样用消息框显示 `WScript.Echo` 和错误 |
| `//UTF8` | 标准输入输出使用 UTF-8（供编辑器读取） |
| `//ErrToOut` | 标准错误写入标准输出（保持两者顺序） |
| `//Nologo` | 兼容参数，不显示版权信息 |
