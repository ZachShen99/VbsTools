# VbsTools 使用说明

VbsTools 为 VS Code 提供完整的 VBScript（`.vbs`）开发环境：能理解 COM 对象的智能提示、错误检查、调试器、兼容 cscript 的脚本宿主，以及代码库。

> 仅支持 Windows。类型信息来自本机注册的 COM 类型库，脚本在插件自带的脚本宿主中运行。

## 目录

| 章节 | 内容 |
|---|---|
| [快速上手](getting-started.md) | 安装、第一个脚本、运行和调试 |
| [编辑与智能提示](editing.md) | COM 补全、悬停、参数提示、错误检查、快速修复、导航、格式化 |
| [运行脚本](running.md) | 运行命令、参数、输出面板与终端、两种宿主、32/64 位、文件编码 |
| [调试](debugging.md) | 断点、单步、变量、监视、调试控制台、异常、launch.json |
| [类型描述与注释](type-files.md) | WinCC 等宿主环境、`*.vbstypes.json`、`@type` 注释、Include 文件 |
| [代码库](code-library.md) | 示例、我的代码片段、快捷键 |
| [设置参考](settings.md) | 全部设置项 |
| [常见问题](faq.md) | 故障排查、与 cscript 的差异 |

## 常用快捷键

| 操作 | 快捷键 |
|---|---|
| 运行脚本 | Ctrl+F5 |
| 带参数运行 | Ctrl+Shift+F5 |
| 停止运行 | Shift+F5 |
| 调试 | F5 |
| 切换断点 | F9 |
| 单步跳过 / 单步进入 / 单步跳出 | F10 / F11 / Shift+F11 |
| 快速修复 | Ctrl+. |
| 转到定义 / 查找所有引用 / 重命名 | F12 / Shift+F12 / F2 |
| 格式化文档 | Shift+Alt+F |

随时可以通过命令面板（Ctrl+Shift+P）执行 **VBScript: 打开说明文档** 回到这里。
