# XMLMCP v0.2.0

新增委托单 XLSX、极片电阻仪原数据 CSV ZIP、报告母版的生成流程。Python 独立计算并画图，再局部注入原模板，Windows 和只有 WPS 的 MacBook 均无需 Microsoft Excel。

- 新工具 `inspect_report_inputs` 和 `generate_report`，共五个 MCP 工具；同一模板首次确认映射后复用，支持不同字段/样品表/结果表/图片位置。
- 逐点均值、样本标准差和 n−1 COV，保留仪器原汇总并记录差异；缺失信息留空并列出待补项。
- 独立散点图/变压曲线 PNG、数值表、局部富文本片段及复选框状态同步；其他对象保持字节。
- 公式、本地/动态图表、透视依赖、容量不足或样品不匹配取消整批。
- 来源与精确补丁完整性报告一并交付；业务文件不包含在公共下载包。

同一源码 ZIP 在 Windows 运行 `python scripts/setup.py`，Mac 运行 `bash scripts/setup.sh`，导入生成的 `mcp.local.json`。[安装说明](https://github.com/Veiluna/XMLMCP/blob/master/INSTALL.md) / [报告生成说明](https://github.com/Veiluna/XMLMCP/blob/master/docs/report-workflow.md)。

自动检查包含原 XML 功能、新统计/映射/失败取消、Mac 分支、内联映射和两种协议的真实 stdio 调用。用户提供模板已在 Windows WPS 12.0 只读核对并导出结果页预览；Mac 真机安装与 WPS 显示仍待复测。通用 `apply_changes` 的公式/透视计算后端继续要求 Windows Excel。
