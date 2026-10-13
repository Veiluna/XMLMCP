# 在其他电脑安装 XMLMCP

适用版本：0.2.2。本地 stdio MCP；Windows 和 macOS 使用同一份源码包，安装脚本在目标电脑创建虚拟环境，并生成该电脑的 `command/args` 配置。

## 支持范围

| 功能 | Windows | MacBook：Apple 芯片 / Intel |
| --- | --- | --- |
| 安装、启动、五个 MCP 工具 | 支持 | 支持；仍待 Mac 真机复测 |
| 委托单 + CSV ZIP + 已确认映射生成报告、统计表和图片 | 支持，无需 Excel | 支持，无需 Excel，可用 WPS 查看 |
| 参考图替换、明确映射的示例图删除、原单元格样式核验 | 支持，无需 Excel | 支持，无需 Excel |
| 制作适配极片母版的可编辑结果表格 | 支持，另行运行准备入口 | 支持，另行运行准备入口 |
| 检查工作表、命名区域、样式及现有对象 | 支持，无需 Excel | 支持，无需 Excel |
| 无需重算的文字、数值、布尔、矩阵、清空、复制已有样式 | 支持，无需 Excel | 支持，无需 Excel |
| 核验部件字节和修改片段 | 支持，无需 Excel | 支持，无需 Excel |
| 公式计算、图表和透视缓存更新 | 需要 Microsoft Excel 桌面版 | 本版未提供 |
| 通用 apply_changes 添加 / 替换图片 | 需要 Microsoft Excel 桌面版 | 本版未提供；报告工具可生成并插入图片 |
| 原有图片、批注、VML、控件等未修改部件的保留 | 支持 | 支持 |

**通用 `apply_changes` 按整个工作簿判断重算需求：** 工作簿只要含有公式、图表、透视表或 chartEx，写入和样式修改就会进入计算流程，即使修改的是其他单元格。此类操作在 Mac 上会返回 `EXCEL_REQUIRED` 或具体的不支持原因，并取消整批，不发布文件。Mac 上添加普通公式也会被拒绝。安装 Excel for Mac 或 WPS 不会启用本版的 Windows Excel 后端。`inspect_template` 的 `capabilities.runtime` 给出实际平台能力；`calculation_supported` 仅表示模板结构是否被适配。

报告工具 `generate_report` 使用单独的离线路径，由 Python 计算原始测量值并画图。模板不可含公式或透视依赖；已验证独立的外部图表快照可原字节保留。它可以在 Mac 上添加报告图片，具体限制见 [报告生成说明](docs/report-workflow.md)。

安装自检使用 `examples/core-template.xlsx`：它含已有图片、文字框、批注、复选框和 VML，没有公式、图表或透视表。其数值是静态示例，不是自动计算结果。

## 准备

1. 从 [GitHub Releases](https://github.com/Veiluna/XMLMCP/releases) 下载 `XMLMCP-0.2.2-portable.zip` 并解压。进入解压后的 **XMLMCP** 文件夹，应能看到 `pyproject.toml`、`src` 和 `scripts`。
2. 安装 [Python 3.12 或更高版本](https://www.python.org/downloads/)。本机已验证 Python 3.14.7；Mac 依赖预检使用 Python 3.12。Windows 安装时让 Python 命令可用；Mac 使用对应 Apple 芯片或 Intel 的官方安装包。
3. 首次安装需要网络下载固定版本依赖。检查和无需重算的修改不要求安装 Office。

**只复制压缩包中的源码，不复制其他电脑的 `.venv`。** Python 虚拟环境包含本机路径，需要在目标位置重新创建。[Python venv 文档](https://docs.python.org/3/library/venv.html)

## Windows 安装

在 PowerShell 中进入实际解压目录。以下 `C:\Tools\XMLMCP` 仅为示例：

```powershell
cd "C:\Tools\XMLMCP"
python --version
python scripts/setup.py
```

如 `python` 不可用，但已安装 Python 启动器，可用 `py -3 scripts/setup.py`。如果默认 Python 低于 3.12，请指定已安装的新版本，例如 `py -3.12 scripts/setup.py`。

脚本创建 `.venv\Scripts\python.exe`，安装锁定依赖，执行 `pip check`，再用官方 SDK 在两种协议路径下实际调用检查、生成、完整性验证三个工具。看到 `core_smoke: passed` 并输出配置文件位置后，安装完成。

也可运行 `scripts/setup.ps1`；若 PowerShell 策略阻止脚本，直接使用上面的 Python 命令即可，无需修改系统策略。

## MacBook 安装

在“终端”中进入实际解压目录。以下目录仅为示例：

```bash
cd "$HOME/Tools/XMLMCP"
python3 --version
bash scripts/setup.sh
```

脚本创建 `.venv/bin/python`，跳过仅供 Windows 使用的 `pywin32`，安装其余锁定依赖并运行同一套 stdio 自检。也可以直接运行 `python3 scripts/setup.py`。

若系统默认 `python3` 太旧，使用已安装的新解释器，例如：

```bash
python3.12 scripts/setup.py
```

Apple 芯片和 Intel 均从源码安装；依赖由 pip 选择对应架构，Intel Mac 的 cryptography 单独固定为仍提供 universal2 安装包的 48.0.1；Apple 芯片和 Windows 使用 50.0.2。[上游变更记录](https://github.com/pyca/cryptography/blob/main/CHANGELOG.rst) 无需 Rosetta 来运行 XMLMCP 本身。依赖预检覆盖 Apple 芯片 macOS 11+、Intel macOS 12+、Python 3.12。实际硬件安装和客户端连接仍需在目标 Mac 完成确认。

## 接入豆包工作或其他本机 MCP 客户端

安装完成后，项目根目录会自动生成 **`mcp.local.json`**，包含目标电脑的绝对路径。导入这个文件；不要复用另一台电脑的配置。

若客户端只提供两个输入框，打开 `mcp.local.json`，复制 `mcpServers.xmlmcp.command` 和 `args`：

| 项目 | 内容 |
| --- | --- |
| command | Windows 为本项目 `.venv\Scripts\python.exe` 的绝对路径；Mac 为 `.venv/bin/python` 的绝对路径 |
| args | `["-m", "xmlmcp"]` |

`command` 的值填写真实路径，无需在 JSON 字符串里面再套引号。Mac 的 `command` 必须保留虚拟环境内的路径，不能改成该软链接指向的系统 Python。

客户端应列出 `inspect_template`、`apply_changes`、`verify_integrity`、`inspect_report_inputs`、`generate_report` 五个工具。第一次先使用核心示例确认连接。报告流程和新模板的映射配置见 [报告生成说明](docs/report-workflow.md)。MCP 文件参数使用目标电脑绝对路径；报告工具创建整个新目录，通用修改工具创建新文件。

直接执行 `python -m xmlmcp` 后等待 stdin 是正常的 stdio 服务行为；诊断用的 `inspect/apply/verify` 参数不填进 MCP 的 args。配置不依赖客户端工作目录。[官方 SDK 接入说明](https://github.com/modelcontextprotocol/python-sdk/blob/main/docs/get-started/real-host.md)

## 安装后的复查与示例

在项目目录中再次执行：

```powershell
# Windows
python scripts/setup.py --check-only
.\.venv\Scripts\python.exe -m xmlmcp --version
```

```bash
# Mac
python3 scripts/setup.py --check-only
.venv/bin/python -m xmlmcp --version
```

`--check-only` 重新运行依赖和 stdio 检查，刷新本机配置，不重新下载安装依赖。自检创建自己的临时结果并清理，不覆盖示例或业务文件。

如果要保留一份可查看的核心示例输出，在项目目录运行以下对应命令。`my-core-report.xlsx` 及同名完整性报告须不存在：

```powershell
# Windows
$xmlmcpRoot = (Get-Location).Path
.\.venv\Scripts\python.exe -m xmlmcp apply "$xmlmcpRoot\examples\core-template.xlsx" "$xmlmcpRoot\examples\my-core-report.xlsx" "$xmlmcpRoot\examples\core-operations.json"
```

```bash
# Mac
xmlmcp_root="$(pwd)"
.venv/bin/python -m xmlmcp apply "$xmlmcp_root/examples/core-template.xlsx" "$xmlmcp_root/examples/my-core-report.xlsx" "$xmlmcp_root/examples/core-operations.json"
```

Windows Excel 的完整示例使用 `examples/template.xlsx` 和安装时生成的 `examples/operations.local.json`，后者含本机图片路径。压缩包中的 `operations.json` 使用占位相对图片路径，不能直接作为完整示例的绝对路径操作输入。

## 移动、升级和排错

- **换电脑或移动文件夹：** 重新解压源码包并安装，导入新生成的 `mcp.local.json`。不要移动虚拟环境后只改配置路径。
- **升级：** 解压到一个新目录，运行安装和自检，通过后将客户端切换到新配置；业务模板和输出不放进发行包。
- **找不到 Python / 版本太旧：** 检查版本，重新打开终端或指定已安装的 Python 3.12+ 解释器。
- **`.venv` 不属于当前平台或不完整：** 使用干净的新解压目录安装；安装脚本不会替你删除旧目录。
- **`EXCEL_REQUIRED`：** 当前操作需要 Windows Excel。Mac 的首版能力见上表；Windows 需安装并能正常启动 Microsoft Excel 桌面版。
- **输出已存在：** 换一个新的输出文件名。MCP 不覆盖模板、结果或完整性报告。

验证记录见 [VALIDATION.md](VALIDATION.md)。Windows 已进行全新解压安装；Mac 的平台分支、依赖和路径配置有自动验证，不能代替 Mac 真机安装验收。
