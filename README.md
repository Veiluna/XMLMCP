# XMLMCP

本机 XLSX 模板注入 MCP。最终文件从原始 ZIP 包生成；未修改部件的解压字节不变，修改部件使用可重放的字节补丁。0.2.0 新增“委托单 Excel + 仪器原始数据 ZIP + 报告模板”生成流程，Windows 和只有 WPS 的 MacBook 均无需 Microsoft Excel。通用公式/透视计算仍使用 Windows Excel 临时副本。

版本下载：[v0.2.2 安装包](https://github.com/Veiluna/XMLMCP/releases/download/v0.2.2/XMLMCP-0.2.2-portable.zip) · [GitHub Releases](https://github.com/Veiluna/XMLMCP/releases)。

## 生成检测报告

先调用 `inspect_report_inputs` 检查字段、单位、样品对应和模板映射建议，首次确认后调用 `generate_report`。同一模板复用映射；更换模板时重新配置字段、样品表、统计表和图片区域。首版适配极片电阻仪单点/变压 CSV ZIP，其他检测项目需增加数据适配器。

输出含 XLSX 报告、散点图和变压曲线 PNG、精确补丁完整性报告、来源/统计 JSON、待补字段及可复用映射。Python 计算并生成静态图片，WPS 直接显示已填结果；源数据变化后重新生成。详见 [报告生成说明](docs/report-workflow.md)。

0.2.2 增加参考图替换和可填表格母版制作：极片单点结果可生成电阻、电阻率、电导率、厚度的均值/COV 四图，变压结果生成电阻率和厚度两图。映射明确指定要替换或删除的示例对象，清理其独占的旧媒体；图框、位置和其他对象保留。图名、坐标轴和表头使用中文，标题字体随包提供。

报告工具只填数据，不复制或新建样式；独立验证原单元格样式 ID、字符格式、样式/主题部件和行列布局。原表格若是 EMF/WMF 图片，先明确制作一份可编辑的新母版，随后所有报告对这份新母版执行样式保留检查。参考 PNG 的图表外观由绘图参数重建，不声称不同数据或不同字体下像素一致。

## 安装与启动

Windows / macOS，Python 3.12+。报告生成、检查、完整性验证和无需重算的 XML 修改不需要 Office；通用 `apply_changes` 的计算缓存和图片修改需要 Windows Microsoft Excel。安装步骤见 [INSTALL.md](INSTALL.md)。

```powershell
cd "你的解压目录\XMLMCP"
python scripts/setup.py
.\.venv\Scripts\python.exe -m xmlmcp
```

```bash
# MacBook：在解压后的 XMLMCP 目录执行
bash scripts/setup.sh
.venv/bin/python -m xmlmcp
```

安装脚本使用 `requirements.lock` 的固定依赖版本。本项目锁定官方 Python SDK `mcp==2.2.0`，以 SDK 的 `MCPServer` 实现协议接入。日志写入 stderr，stdio 的 stdout 只承载协议。

安装成功会自动生成当前电脑的 **mcp.local.json**。在客户端的本机 MCP 配置入口导入该 JSON；若入口分别填写字段，复制其中 command/args。以下为格式示例，command 使用安装脚本生成的真实绝对路径：

```json
{
  "command": "C:\\Tools\\XMLMCP\\.venv\\Scripts\\python.exe",
  "args": ["-m", "xmlmcp"]
}
```

不依赖客户端工作目录。终端直接启动后等待 stdin 是正常行为。其他电脑使用发行 ZIP 重新安装，不复制 `.venv`；新目录安装后使用新配置。飞书文档中的企业 HTTP 接入与本项目的本机 stdio 接入分开配置。

## 五个工具

- `inspect_template(path, targets=[])`：查看指纹、名称作用域、目标值和样式、图片 ID、图表来源及支持情况。
- `apply_changes(input_path, output_path, operations)`：整批预检、修改、计算、验证，成功后发布新文件及 `.integrity.json` 报告。
- `verify_integrity(input_path, output_path, report_path)`：只读重放报告中的补丁，核验逐部件哈希、片段外字节和包内引用。报告是修改记录，不是第三方签名证明。
- `inspect_report_inputs(template_path, order_path, raw_zip_path, profile_path?, metadata?)`：只读检查样品和字段、独立计算统计，返回映射建议和待补信息。
- `generate_report(template_path, order_path, raw_zip_path, output_directory, profile_path?, metadata?, profile?)`：无需 Excel 生成报告；内联 `profile` 与文件 `profile_path` 二选一，输出目录必须为不存在的新目录。

所有文件路径必须是绝对路径。输入不会覆盖；输出文件和报告必须尚不存在。每次调用完整指定输入及目标。

目标格式：`{"name":"报告标题","scope":"workbook"}` 或 `{"sheet":"报告","range":"B3"}`。局部名称的 scope 为工作表名称。建议每次写入均指定 scope；省略时只接受在整个工作簿中唯一的名称，同名作用域不唯一会拒绝。

操作示例：

```json
[
  {"op":"write","target":{"name":"报告标题","scope":"workbook"},"value":"本期报告"},
  {"op":"write","target":{"sheet":"报告","range":"B8:C10"},"value":[["甲",120],["乙",180],["丙",240]]},
  {"op":"formula","target":{"sheet":"报告","range":"C12"},"formula":"=SUM(C8:C10)"},
  {"op":"copy_style","source":{"sheet":"报告","range":"B8"},"target":{"sheet":"报告","range":"B9:B10"}},
  {"op":"add_image","target":{"name":"照片区域","scope":"workbook"},"image_path":"G:\\资料\\照片.png"}
]
```

`write` 中以 `=` 开头的字符串仍是文字；`null` 和 `clear` 清空值。矩阵不能超出预留范围，剩余单元格不改。标量必须对应单格或完整合并单元格；合并区域只写左上角。`replace_image` 使用检查工具返回的 `image_id`。重叠的值写入会拒绝，样式复制与值写入可同时作用于同一单元格。

日期默认写文字。Excel 日期值须写入序列值，并复制模板已有日期样式。超过 15 位有效整数精度的编号请写文字；单格文字最多 32767 个 UTF-16 单位。回车、字面量 `_xHHHH_`、空格及 XML 特殊字符均有专门处理。

## 可运行的示例

[examples/core-template.xlsx](examples/core-template.xlsx) 与 [examples/core-operations.json](examples/core-operations.json) 可在两端运行：含已有图片、文字框、批注、复选框和 VML，无公式、图表或透视表。安装时会自动在两种协议下生成、验证临时结果。保留示例输出的命令见 [INSTALL.md](INSTALL.md)。

Windows Excel 完整示例 [examples/template.xlsx](examples/template.xlsx) 由 Excel 创建，含图片、图表、复选框和批注。本工作区还保留了已生成的 result.xlsx 和完整性报告；发行 ZIP 不携带这些机器相关结果。安装脚本会根据当前目录生成 `examples/operations.local.json`；使用它生成另一份结果（`my-report.xlsx` 必须不存在）：

```powershell
$xmlmcpRoot = (Get-Location).Path
.\.venv\Scripts\python.exe -m xmlmcp inspect "$xmlmcpRoot\examples\template.xlsx"
.\.venv\Scripts\python.exe -m xmlmcp apply "$xmlmcpRoot\examples\template.xlsx" "$xmlmcpRoot\examples\my-report.xlsx" "$xmlmcpRoot\examples\operations.local.json"
.\.venv\Scripts\python.exe -m xmlmcp verify "$xmlmcpRoot\examples\template.xlsx" "$xmlmcpRoot\examples\my-report.xlsx" "$xmlmcpRoot\examples\my-report.integrity.json"
```

MCP 调用参数与这些 CLI 参数对应；CLI 仅用于诊断，豆包配置中保持 `args: ["-m", "xmlmcp"]`。

## 边界

支持未加密、未签名的 Transitional `.xlsx`；数据注入不插删或移动行列，不创建任意样式。`prepare-report-template` 是另行执行的母版制作入口，只适配说明中的已标记极片母版，在新输出中创建表格样式；报告工具不会隐式调用它。文本用 inlineStr，保留已有 s 和一致的富文本字符格式。混合富文本和拼音/语音标注拒绝整格覆盖。只存于字符运行中的字体不能在数值、公式或清空操作中可靠保留，此类类型转换也会拒绝。

Mac 支持检查、无需重算的 XML 修改及 `generate_report` 的 Python 计算、画图和 XML 注入。通用 `apply_changes` 仍以整个工作簿为计算单位，要求 Windows Excel 后端。报告工具使用严格的离线路径：无工作簿公式/透视表，原有图表只能是可证明与新数据无关的静态文字或外部单元格快照；本地、动态或未知来源会拒绝。`capabilities.runtime` 报告实际平台能力。

共享、数组和数据表公式区域不能直接写入；已有固定数组的依赖数据可改，缓存逐格回填，结果范围必须固定。新增公式必须能作为普通公式计算并保持表达式语义，Excel 改写为未适配函数/数组、产生新错误或无法映射时整批失败。

图表支持本机 Excel 能计算且可唯一映射、缓存可交叉验证的标准图表来源，包括单元格、公式、动态名称、固定布局透视表及外部快照。透视结果范围、字段或项目布局变化会拒绝。多层分类图表、chartEx、分组透视和必要的未知二进制数据模型更新尚不支持。

图片支持 PNG/JPEG，保持比例、居中；替换保留锚点和对象属性，重置裁剪，使用新媒体关系并保留共享原图。复杂图片效果、分组、颜色变换、旋转、EXIF 方向及无法唯一映射的对象不支持替换。这些对象在不修改时原字节保留。

外部数据须事先在 Office 准备，MCP 不联网刷新外部源。已保存快照可用于计算/显示；需要外部实时计算或未适配数据模型的操作会失败。

## 验证

当前结果见 [VALIDATION.md](VALIDATION.md)。报告测试覆盖独立统计、编码/NUL 处理、不同模板映射、来源不变、容量/样品/依赖失败取消，以及两种协议的新工具真实调用。原有 Excel 计算验收记录按版本保留。

```powershell
.\.venv\Scripts\python.exe -m pytest -q --basetemp=test-artifacts\unit-portable --junitxml=validation\portable-unit-tests.xml
.\.venv\Scripts\python.exe scripts\check_install.py
.\.venv\Scripts\python.exe scripts\protocol_smoke.py examples\template.xlsx --apply
.\.venv\Scripts\python.exe scripts\office_smoke.py --kind all --wps
```

Mac 用 `.venv/bin/python` 运行 pytest 或 check_install.py；Office 脚本仅供 Windows。用户提供模板的报告已通过字节检查和 Windows WPS 首次打开核对；Mac 真机及人工提示检查仍待复测。SDK 测试不代替豆包实测，导入本机配置后确认五个工具可见。

文档依据：[飞书 MCP 使用](https://aily.feishu.cn/hc/1u7kleqg/73ahi33a)、[官方 SDK](https://github.com/modelcontextprotocol/python-sdk)、[MCP 开发文档](https://modelcontextprotocol.io/docs/2026-07-28/develop/build-server)。字符串转义依据 [Microsoft ST_Xstring 说明](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/bd0aa042-434a-4ca7-b25f-4e1fd25a954d)。实现结构见 [docs/architecture.md](docs/architecture.md)。
