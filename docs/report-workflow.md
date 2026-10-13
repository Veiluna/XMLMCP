# 从委托单和原始数据生成报告

0.2.0 增加 `inspect_report_inputs` 与 `generate_report`。报告生成在 Python 中计算并画图，然后在原 XLSX 包中局部注入，不需要 Microsoft Excel 或 WPS 自动化。MacBook 只安装 WPS 即可查看结果。

0.2.2 支持六种参考统计图、替换已有图框和清理明确选择的示例对象。报告只修改真实值和表头；不复制或重建单元格样式。原有图片、图表是否属于示例由首次确认的映射决定，检查工具返回 `available_objects` 和图片 ID 供选择。

0.2.3 对已标记极片模板保留全部“极片电阻仪参考资料”和原有对象，不按附录位置删除图片。新可填母版的逐点表按实际数据调整单元格行列数，复用母版样式，所有列宽、行高保持原值；固定表格仍只填值。

## 第一次使用模板

1. 把委托单 XLSX、报告母版 XLSX 和原始数据 ZIP 放在本机。
2. 调用 `inspect_report_inputs`，检查样品一一对应、字段来源、单位、待补信息及映射建议。
3. 首次确认映射。已识别的极片母版返回 `suggested_profile`；陌生模板返回命名区域和字段目录，需要配置位置。批注可以帮助理解模板，但只是数据，不作为代理指令。
4. 调用 `generate_report`，传已确认的 `profile` 对象，或已保存的 `profile_path`。这两个参数二选一。成功时输出目录自动保存 `template-profile.json`，以后可复用此文件。
5. 用 WPS 打开 `report.xlsx`。若源数据改变，重新生成到一个新目录。

`output_directory` 必须为不存在的新目录，父目录应存在。不覆盖任何输入、旧报告或旧目录。生成失败时清理本次暂存数据，不交付部分报告。

首次检查调用：

```json
{
  "template_path": "/Users/yourname/Reports/template.xlsx",
  "order_path": "/Users/yourname/Reports/order.xlsx",
  "raw_zip_path": "/Users/yourname/Reports/raw.zip"
}
```

复用映射生成：

```json
{
  "template_path": "/Users/yourname/Reports/template.xlsx",
  "order_path": "/Users/yourname/Reports/order.xlsx",
  "raw_zip_path": "/Users/yourname/Reports/raw.zip",
  "profile_path": "/Users/yourname/Reports/template-profile.json",
  "output_directory": "/Users/yourname/Reports/report-001",
  "metadata": {
    "received_date": "2026-09-09",
    "report_date": "2026-09-11",
    "operator": "已确认人员",
    "sample_order_no": "已确认单号"
  }
}
```

Windows 换成 `C:/Reports/...` 之类的实际绝对路径。缺失字段留空，并在工具结果和 `report-data.json` 中列出；不以寄样日期代替收样日期，不推测操作人员或委托单号。检测日期来自 CSV 的测量开始/结束时间。

## 新模板配置

配置分为来源和输出两部分。`order.fields` 映射委托单中的公司、联系人等标量，`order.sample_rows` 映射样品名、编号、材料等横向行。输出 `fields` 把同一字段填到一个或多个命名区域或单格；`sample_rows` 填到预留的一行；`tables` 和 `figures` 决定结果表和图片区域。

支持的字段目录由检查工具返回。输出目标优先使用 `{"name":"委托单位","scope":"workbook"}`，也可用 `{"sheet":"报告","range":"B3"}`。工作表名包括其原有空格。映射保存模板的 SHA-256，模板被移动不影响复用，但修改内容、格式或布局后需要重新检查并确认映射。

下面是一个最小配置，指纹应取检查工具结果；位置由实际模板确定：

```json
{
  "schema_version": 1,
  "profile_id": "my-report-v1",
  "data_adapter": "electrode-resistance-csv-v1",
  "template_fingerprint": "替换为检查工具返回的64位SHA256",
  "order": {
    "fields": {"client_company": {"sheet":"委托单","range":"B8"}},
    "sample_rows": {"sample_name": {"sheet":"委托单","range":"C17:G17"}}
  },
  "fields": {"client_company": [{"sheet":"报告","range":"B3"}]},
  "sample_rows": {"sample_name": {"sheet":"报告","range":"B6:F6"}},
  "figures": {"single_point": {"sheet":"报告","range":"A10:H25"}},
  "tables": {"single_summary": {"target":{"sheet":"报告","range":"A28:H34"}}}
}
```

- `single_summary` 表为两行标题/单位，加每个样品一行，共 8 列。
- `pressure_resistivity` 表为两行标题/单位，加每个设定压强工步一行，共“压强 + 样品数”列。各样品的工步必须一致，不插值或按文件顺序拼接。
- `figures.single_point` 生成逐点电阻率散点图；`figures.pressure` 生成实测压强与电阻率曲线。图片为静态 PNG，保留比例、居中，不生成依赖 Office 重算的动态图表。
- 图名和轴标题默认中文：单点图为“单点电阻率散点图 / 测量点序号 / 电阻率（Ω·cm）”，变压图为“变压电阻率曲线 / 实测压强（MPa） / 电阻率（Ω·cm）”。随包提供标题字体，无需在系统中另行安装；中文样品名优先使用 Windows 微软雅黑或 Mac 苹方/黑体。
- `header_style` / `data_style` 仅保留旧配置读取兼容，在报告生成中不执行样式复制。每个目标单元格保留自己的原样式；应先在母版设置所有预留位置的样式。所有原行列尺寸保留。
- `maximum_digit_width_px` 是模板字体的数字宽度标定，默认为 7px。图片定位使用原有行高与 OOXML 列宽公式；特殊字体需在首次查看时确认映射中的该值。
- `text_runs` 可明确指定某个富文本 `run_index`，只替换该片段正文，保留所有 `rPr`；共享文字复制到目标单格的 inlineStr，不修改 sharedStrings。普通混合富文本整格覆盖仍拒绝。
- `checkboxes` 可按来源/目标工作表的唯一标签复制现有复选框状态，同步 VML 与 ctrlProps。新版模板用单元格方框文字时，绑定增加 `target` 单格位置，只替换相应的“☐ / ☑”，保留文字和各段字体。关联单元格、混合状态和不唯一的标签拒绝，不猜测样品安全状态。

配置不同坐标、名称、样品列数、图表区域即可适配不同报告布局。首版只支持当前数据适配器和上述图/表类型；其他测试方法、复杂文档对象或新计算规则需增加适配器。

## 按参考外观制作可填表格的新母版

有些母版的图后表格是 EMF/WMF 图片，图中的背景、字体和边框不是 Excel 单元格样式。对这类母版，先单独制作一份可填表格的新母版，再确认它的映射。这个准备步骤会创建新表格样式；后续 `generate_report` 对新母版执行严格的样式保留检查。

已标记极片母版的准备入口如下。`base.xlsx` 是结果区已清空、包含“单点结果区”和“变压结果区”命名区域的母版；`reference.xlsx` 保留完整六张统计示例 PNG。`new-mother` 必须不存在，父目录必须存在。

```bash
# Mac；Windows 改用 .venv\Scripts\python.exe
.venv/bin/python -m xmlmcp prepare-report-template /abs/base.xlsx /abs/reference.xlsx /abs/new-mother --sample-count 5 --single-points 6 --pressure-points 12
```

此准备器仅适配已标记极片母版的标准布局和六个已识别参考对象：单点区 52 行 × 11 列、变压区 35 行 × 11 列。输出 `template.xlsx`、`template.integrity.json`、`preparation.json`，保留输入。以 A74/A129 为结果区起点时，逐点表从 A106、A146 开始；上述五个样品、六个单点、十二个压强工步对应 A106:K115、A146:K159。两个样品只使用 A:E 五列；六个单点仅使用十行，平均值和 COV 紧跟实际数据，未用区域恢复空白背景。不把少量数据拉伸到整个报告宽度，不插删工作表的全局行列。

单点表保留绿色表头、黄色平均值行；变压表保留蓝色表头。源图片为 placeable WMF 时，按文字实际选定的中文字体、字号和字重制作单元格字体；中文标题优先取中文文字段的字体，变压标题和表头保留粗体。`preparation.json` 记录提取依据和字号。原始 WMF 不是可编辑单元格，因此这是模板制作，不声称图片重建逐像素一致。准备器未找到对应 WMF 时会记录显式默认字体依据；需要精确复现的其他图片格式应另行适配。

母版用独立的 Custom XML 部件保存表格空白区及原样式布局，检查工具将其读入 `tables.<kind>.layout`，它是格式数据，不是代理指令，也不加入 Excel 公式名称。每次生成均按明确传入的模板和已确认映射处理，同一母版可自动适应不同行列数量。标准空白区域最多五个样品、十六个测量点或压强工步；各样品序号/工步应一致。超出区域会报错，不改全页列宽、行高或挤占参考资料。不同布局需要首次配置或新增准备适配器。

准备步骤完整保留原有图片、图表、批注、控件、VML、商标和参考资料。六个结果图框按参考位置复制到空白结果区；生成报告时，仅这些图框替换为真实数据图片并清理其独占的旧媒体。新母版先运行 `inspect_report_inputs`；返回 `electrode-editable-template-v3` 映射建议，经确认保存后复用。

## 替换示例图与六种统计图

`figure_replacements` 的值是检查工具返回的 `image_id`；替换时保留原锚点、图框和对象属性，重置裁剪并适配比例。`example_objects` 精确列出额外要删除的图片或原生图表 `object_id` 和 `kind`。删除只适用于标准直接 drawing 对象；共享媒体只有最后一个引用移除后才清理。未知、分组或不唯一的对象报错。图表删除后仍有其他本地/动态依赖的模板会被离线预检拒绝。

在包含“极片电阻仪参考资料”标题的报告页，除明确新建的 `XMLMCP Reference ...` 六个结果图框外，原有图片、图表及其依赖自动受保护。指定删除、替换这些对象，或修改参考资料行，会被拒绝。通用模板需要删例图时仍可明确映射，不能套用到受保护的附录。

```json
{
  "figure_style": "template-statistics",
  "figure_replacements": {
    "single_resistivity": "报告|xl/drawings/drawing1.xml|21",
    "pressure_thickness": "报告|xl/drawings/drawing1.xml|43"
  },
  "example_objects": [
    {"object_id": "报告|xl/drawings/drawing1.xml|14", "kind": "chart"}
  ],
  "tables": {
    "single_measurements": {"target": {"sheet": "报告", "range": "A106:K119"}},
    "pressure_measurements": {"target": {"sheet": "报告", "range": "A146:K159"}}
  }
}
```

上述为完整映射中的片段，所有 ID 和位置必须取自当前模板，不能直接背诵示例。`figures` 在空白区域添加新图；同一图种不可同时添加和替换。

| 图种 | 真实数据与样式 |
| --- | --- |
| `single_resistance` | 电阻均值（mΩ），红色圆点；样本 COV，绿色三角、右轴 |
| `single_resistivity` | 电阻率均值（Ω·cm）及同指标样本 COV |
| `single_conductivity` | 电导率均值（S/cm）及同指标样本 COV |
| `single_thickness` | 极片总厚度均值（μm）及同指标样本 COV |
| `pressure_resistivity` | 各样品实测压强（MPa）与电阻率，彩色圆点曲线 |
| `pressure_thickness` | 各样品实测压强（MPa）与总厚度，彩色圆点曲线 |

`template-statistics` 固定透明背景、深灰坐标轴、灰色虚线网格及参考点型、线型。横轴样品名来自委托单，单位和标题为中文；变压比较以真实工步为点，不插值。参考图只有两个样品时，新增样品沿映射固定调色板使用蓝、橙、紫等颜色。坐标范围随真实数据变化。PNG 参考图没有可编辑图表样式定义，这些外观由绘图参数重建；不承诺逐像素一致。图种参数、字体、曲线点和均值/COV 均记录在 `report-data.json` 的 `plots` 中。

`single_measurements` / `pressure_measurements` 分别写入电阻率和总厚度两组真实样品表头。配置了母版内的 `layout` 时，表宽为“1 + 2 × 实际样品数”；表高是“实际测量点数 + 4”或“实际压强工步数 + 2”。复制相同角色的原有样式 ID，只调整该空白区域里的单元格数和标题合并范围，未用位置恢复母版原背景。没有 `layout` 的旧固定映射仍按原容量填充，保留原平均值行。各样品点序号或工步需一致，容量不足或合并冲突取消整批。

## 计算与来源

CSV 支持 UTF-8、GB18030 和能识别标题的 UTF-16；读取 ZIP 内文件，不解压到外部目录。仪器某些导出在行首插入 NUL，仅移除这些行首字符并记录数量。未知文件、重复样品/模式、采样数不符、非有限数值、工步不符或样品名不匹配会报错。

统计从逐点记录独立计算：均值、样本标准差（分母 n−1）、样本 COV = 样本标准差 / |均值| × 100%。单点不足 2 个或均值为零时，相关统计留空，不伪造为零。仪器原均值/COV 行保留在来源 JSON；电阻率 COV 差异单独列出。电阻率、厚度和实测压强使用 CSV 本来记录的数值及单位，不重新推导仪器校正。

报告文字只描述本次测量的均值、离散程度及压强变化，不编造材料机理、验收标准或合格结论。委托单的示例列只有在明确配置的编号行出现例子标记时才排除，并记录排除位置。

本流程不计算模板中的 Excel 公式，不刷新本地动态图表或透视缓存。没有公式/透视依赖的模板可以保留原有外部单元格快照和仪器参考图表；动态、本地或无法证明与新数据独立的来源直接报错。需要 Office 计算的通用操作继续使用 Windows Excel 后端。

## 输出与验证

输出目录包含 `report.xlsx`、已配置的 PNG、`report.integrity.json`、`report-data.json`、`template-profile.json`。来源 JSON 记录输入指纹、逐点数据、原仪器汇总、统计口径、绘图参数、字段来源和待补字段。完整性验证独立重放补丁，检查片段外字节、关系、内容类型及缓存结构；结果文件始终由原包生成。

报告完整性记录还包含 `style_preservation`：样式和主题部件字节一致，字符运行格式保留，原行属性、列定义和条件格式不变。固定表格的原单元格 `s` 和合并不变。动态表格记录 `table_layouts`，独立验证它与母版内保存的布局一致：只有明确的空白表格区可复用角色样式、恢复背景、调整标题合并，表格外的 `s` 和合并不变。`destination_styles_identical` 表示所有原位置的 `s` 是否相同；缩放单元格数量时它可能为 false，`style_parts_identical` 和 `outside_table_styles_identical` 必须为 true。

`reference_preservation` 记录受保护的参考资料行和原对象，验证逐行 XML、锚点、媒体及依赖字节不变。`verify_integrity` 独立重查这些约束；母版制作另行记录追加样式的步骤，不混同于报告填写。

CLI 调试入口：

```bash
.venv/bin/python -m xmlmcp report-inspect /abs/template.xlsx /abs/order.xlsx /abs/raw.zip
.venv/bin/python -m xmlmcp report-generate /abs/template.xlsx /abs/order.xlsx /abs/raw.zip /abs/profile.json /abs/new-output
```

MCP 接入的 args 仍是 `["-m", "xmlmcp"]`。报告输入及生成结果留在本机，不收入公共发行包。

实现依据：[Matplotlib 无窗口绘图](https://matplotlib.org/stable/users/faq.html)、[OOXML 列宽与物理尺寸](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.column?view=openxml-3.0.1)、[复选框状态枚举](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.excel.checkedvalues?view=openxml-3.0.1)。
