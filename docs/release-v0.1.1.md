# XMLMCP v0.1.1

Windows / MacBook 本机 stdio MCP 安装包，用于检查 XLSX 模板、执行无需重算的 XML 修改，以及核验逐部件字节完整性。

## 下载与安装

下载本 Release 的 `XMLMCP-0.1.1-portable.zip`，解压进入 XMLMCP 文件夹，先安装 Python 3.12+。

- Windows PowerShell：`python scripts/setup.py`
- Mac 终端：`bash scripts/setup.sh`

安装成功后，将自动生成的 `mcp.local.json` 导入支持本机 stdio 的 MCP 客户端。安装脚本会检查依赖，并在新旧两种协议路径下实际调用三个工具。不要复制另一台电脑的 `.venv`。

## 支持范围

- Windows / Mac：模板检查；无需重算的文字、数值、布尔、矩阵、清空与已有样式复制；完整性验证。
- Windows Microsoft Excel：公式计算、图表与透视缓存更新、图片添加与替换。
- Mac 首版未提供重算和图片修改。当前按整个工作簿判断计算依赖；含公式、图表或透视表的模板修改会被拒绝并取消整批。
- Intel Mac 的 cryptography 固定为兼容的 48.0.1；Apple 芯片和 Windows 使用 50.0.2。

## 验证记录

Windows 全新安装与 46 项自动测试通过；Apple 芯片 / Intel Mac 的目标依赖 wheel 和版本约束检查通过。Mac 真机、豆包原生界面连接及真实脱敏业务原模板尚待验收。发布工作流还会独立运行无需 Office 的 stdio 自检和自动测试。

压缩包旁的 `.sha256` 文件用于核验下载完整性。安装步骤与完整边界见仓库的 INSTALL.md 和 VALIDATION.md。
