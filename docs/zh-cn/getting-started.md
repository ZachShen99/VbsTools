# 快速上手

## 安装

1. 在 VS Code 中打开扩展视图（Ctrl+Shift+X），点右上角 **…** → **从 VSIX 安装…**，选择 `vbstools-vbscript-*.vsix`。
   也可以在命令行执行：

   ```powershell
   code --install-extension .\vbstools-vbscript-win32-x64-0.1.1.vsix
   ```

2. 安装后执行 **Developer: Reload Window**（Ctrl+Shift+P）。
3. 升级时如果版本号没变，加 `--force` 覆盖安装。

插件自带语言服务器和脚本宿主（x64 / x86），不需要另外安装 .NET。

## 第一个脚本

1. 新建文件，保存为 `hello.vbs`。右下角的语言模式应显示 **VBScript**。
   如果显示的是 "Visual Basic"，见[常见问题](faq.md#vbs-文件显示为-visual-basic)。
2. 输入：

   ```vbscript
   Option Explicit

   Dim fso, folder
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set folder = fso.GetFolder("C:\Windows")
   WScript.Echo folder.Files.Count & " files"
   ```

   输入 `fso.` 时会列出 FileSystemObject 的所有成员；输入 `CreateObject("` 时会列出本机注册的 ProgID。
3. 按 **Ctrl+F5** 运行。结果显示在 "VBScript" 输出面板，出错时错误会显示在"问题"面板，并跳到出错的行。
4. 在 `WScript.Echo` 那一行按 **F9** 设断点，按 **F5** 调试。停下后在"变量"视图里可以展开 `folder` 查看它的属性。

## 状态栏

打开 `.vbs` 文件时，右下角显示 **VBScript x64** 或 **VBScript x86**。点一下可以切换 64 位 / 32 位，只注册了 32 位的 COM 组件（部分旧版 Office 插件、数据库驱动）需要 x86。

## 下一步

- [编辑与智能提示](editing.md)：补全、错误检查和快速修复的详细说明
- [运行脚本](running.md)：参数、两种宿主、文件编码
- [调试](debugging.md)
- 脚本运行在 WinCC 等软件内部？见[类型描述与注释](type-files.md)
