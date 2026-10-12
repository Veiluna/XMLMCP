# XMLMCP 验收记录

2026-10-12，本机 Python 3.14.7、MCP SDK 2.2.0、Excel 16.0、WPS 12.0。原始 Office 验收属于 0.1.0，汇总证据和当时源码指纹保存在开发工作区 validation/acceptance.json。0.1.1 的安装与跨平台回归证据保存在 [validation/portable-install.json](validation/portable-install.json)，发行 ZIP 包含后者与两种 Mac 架构的 wheel 预检记录。

## 0.1.1 跨平台安装与回归

| 验收项 | 结果与证据 |
| --- | --- |
| 自动测试 | 46 通过；包含原 35 项，以及平台依赖、Mac 基础 XML 操作与对象保留、需 Excel 操作的提前拒绝、配置迁移、POSIX 发布与发行包边界 |
| Windows 全新安装 | 从源码 ZIP 解压到含中文和空格的新目录，创建新的虚拟环境，安装、pip check 和两种协议下的三个工具自检全部通过；没有复制旧虚拟环境 |
| 当前工作区安装 | 用更新后的平台锁文件重新安装，生成当前目录的 mcp.local.json；同样通过依赖和 stdio 检查 |
| Mac Apple 芯片依赖 | macOS 11+、Python 3.12 的 35 个锁定包均有匹配 wheel；52 条有效依赖版本约束通过 Darwin 标记检查 |
| Mac Intel 依赖 | macOS 12+、Python 3.12 的 35 个锁定包均有匹配 wheel；cryptography 为 48.0.1，其余版本与 Apple 芯片相同；52 条有效依赖版本约束通过 |
| 安装配置 | Windows Scripts/python.exe 与 POSIX bin/python 路径分别生成；command 不解析 POSIX 虚拟环境软链接；图片示例改为目标电脑路径 |
| 发布 | 原有输出不得覆盖；POSIX 发布分支的硬链接和竞争失败已在 Windows 文件系统测试，尚未在真实 APFS 上测试 |

Mac wheel 预检使用 `pip --dry-run --no-deps` 指定目标架构；依赖约束另外使用显式 Darwin / Python 3.12 标记核验。Windows 上的 pip 仅指定 `--platform` 不会改变宿主系统标记，不能将其当作真正的 Mac 安装。SDK、XML 引擎和安装脚本的运行测试在 Windows 进行；Mac 真机安装、APFS 发布和客户端连接仍待验收。

## 0.1.0 原生 Office 已通过

| 验收项 | 结果与证据 |
| --- | --- |
| 自动测试 | 35 通过，0 失败；名称作用域、矩阵/空值/合并区域、新单元格、样式/富文本、图片与共享媒体、越界/冲突、发布竞争与字节篡改等 |
| stdio 接入 | 官方客户端分别以 2025-11-25 与 2026-07-28 协议启动项目虚拟环境 Python，从另一个工作目录调用检查、生成和验证三个工具；两份生成结果均通过完整性检查 |
| 字节保留 | 无计算依赖的文字修改只改目标 worksheet；其他所有原部件解压字节一致。图片、计算依赖和透视修改逐补丁重放，片段外相同 |
| 普通图表 | 新数据 100/200；Excel、WPS 首次打开读取一致，导出图中显示相应柱高 |
| 动态名称图表 | OFFSET 名称来源读取到 100/200，不将动态名称作为写入目标 |
| 透视图表 | 固定布局数据更新为 10/20/30；透视缓存、记录和输出网格均回填，Excel/WPS 一致 |
| 外部快照图表 | 保持事先准备的 7/9；工作进程没有外部链接更新或 RefreshAll 调用 |
| 公式缓存 | 新增/修改公式的数值 600/200、布尔 true、文字“金额100”均在 ZIP 内具有正确缓存，Excel/WPS 首次读取一致 |
| 固定数组 | 原数组 E5:E6 更新为 200/400，公式与范围保持，两个客户端读取一致 |
| 字符格式与对象 | 7 磅 Arial 英文、批注、勾选状态、文字说明和图片对象保持；图片添加与替换保留指定框架几何尺寸 |
| 文字编码 | `=SUM(1,2)` 为文字；字面量 `_x0041_`、XML 特殊字符、回车/换行/Tab 在 Excel/WPS 保持输入内容 |
| 失败取消 | 除零错误、函数/数组无法映射、透视布局变化、透视输出被计算覆盖、重复图片 ID 均拒绝；相关测试确认不发布文件、输入不变 |
| 安装检查 | setup.ps1 执行成功，固定版本依赖安装完成，pip check 与源码编译检查通过 |

原始 Office 验收证据在 `test-artifacts/office/ae81f5bb83/acceptance.json`；公式与失败回归在 `test-artifacts/office/regression-699860fd35/acceptance.json`；两种协议的真实调用记录在 `test-artifacts/office/ae81f5bb83/protocol-44f45d41c7/acceptance.json`。汇总文件记录这些绝对路径和图表图片哈希。WPS Chart.Export 的布尔返回值与文件生成不一致，汇总检查实际 PNG 文件及其有效性。

## 尚待完成

- **豆包工作真实连接：** 配置已生成，精确 command/args 已通过 SDK 实测。现有工具无法控制豆包工作的原生配置界面，未将 SDK 测试记为豆包实测。导入安装时生成的 mcp.local.json 后，应在客户端确认三个工具可见并使用核心示例模板生成一个新文件。
- **Mac 真机：** 安装脚本、目标架构依赖和平台分支有验证；真实安装、首次启动、APFS 输出发布和豆包工作连接未在 Mac 上实测。
- **真实原模板：** 用户授权创建项目目录，没有指定一份原始脱敏业务 XLSX。工作区后来出现的 openpyxl 兼容版本仅做了只读检查，没有当作原始无损验收基线，也没有修改它。
- **交互提示：** 已在隔离隐藏实例中完成 Excel/WPS 首次打开、缓存和对象检查，没有 F9 或二次保存。该 API 检查不等价于人工观察“无修复提示”；真实模板需要再完成这一项。

## 复测

从 README 的测试命令运行，Office 脚本会创建新的独立目录，不覆盖已有结果。要复测公式/失败场景，将成功的 Office 目录传给 `scripts/office_regression.py <目录> --wps`。示例模板、操作列表、结果和完整性报告位于 [examples](examples/operations.json)。
