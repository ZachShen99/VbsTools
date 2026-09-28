# 设置参考

在设置中搜索 `vbscript` 可以看到全部设置。

## 编辑

| 设置 | 默认值 | 说明 |
|---|---|---|
| `vbscript.format.keywordCase` | `pascal` | 格式化时是否统一关键字和名称的大小写；`preserve` 只调整缩进 |
| `vbscript.inlayHints.parameterNames` | `literals` | 参数名提示：`none` 不显示；`literals` 只在字面量参数前显示；`all` 在每个参数前显示 |
| `vbscript.codeLens.references` | `true` | 在过程和类上方显示引用数 |
| `vbscript.environments` | `[]` | 所有脚本都启用的宿主环境，例如 `["WinCC"]` |
| `vbscript.typeFiles` | `[]` | 工作区之外的类型描述文件或文件夹 |

## 运行和调试

| 设置 | 默认值 | 说明 |
|---|---|---|
| `vbscript.run.host` | `vbstools` | 运行脚本的宿主：`vbstools` 或 `cscript`；调试始终使用 VbsTools 宿主 |
| `vbscript.run.platform` | `x64` | `x64` 或 `x86`（只注册了 32 位的 COM 组件） |
| `vbscript.run.console` | `auto` | 输出位置：`auto`（用到 StdIn 时用终端）、`output`、`terminal`；也是调试时的默认值 |
| `vbscript.run.scriptHostPath` | 空 | 自定义 VbsTools 宿主所在的文件夹（包含 `x64\` 和 `x86\`） |

## 文件编码

| 设置 | 默认值 | 说明 |
|---|---|---|
| `vbscript.encoding.ansi` | `auto` | ANSI 脚本使用的代码页；`auto` 为系统 ANSI 代码页 |
| `vbscript.encoding.check` | `true` | 编码会让 cscript 读错时显示警告 |

## 代码库

| 设置 | 默认值 | 说明 |
|---|---|---|
| `vbscript.library.samplesFile` | 空 | VbsEdit 示例文件的位置；空表示在 VbsEdit 安装目录中查找 |

## 排查问题

| 设置 | 默认值 | 说明 |
|---|---|---|
| `vbscript.languageServer.path` | 空 | 自定义语言服务器的路径 |
| `vbscript.languageServer.trace` | `false` | 在 "VBScript Language Server" 输出通道中记录每次补全请求（修改后需重启语言服务器） |
| `vbscript.debug.trace` | `false` | 在 "VBScript Debug Trace" 输出通道中记录调试器的全部通信 |

## 其他相关设置

- `"files.associations": { "*.vbs": "vbscript" }`：`.vbs` 被识别为 Visual Basic 时使用
- `editor.inlayHints.enabled`：VS Code 自带的开关，也控制参数名提示
