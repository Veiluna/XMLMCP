# XMLMCP v0.2.2

报告生成现在可替换模板的示例图，按真实数据生成四张单点均值/COV 图和两张变压比较图；通过明确映射删除附录示例对象，并仅清理不再共享的媒体和图表依赖。

表格只填写真实值和真实表头，保留每个目标单元格的底色、字体、边框、数值格式及原行列布局。完整性报告和独立验证增加样式保留检查，旧映射的 `header_style` / `data_style` 不再复制样式。

新增 `prepare-report-template` CLI，为已标记极片母版把图片表格制作成可填的原生单元格表格。这是明确执行的母版制作步骤；以后以新母版为基准只填数据。参考 PNG 图表的外观由绘图参数重建，不承诺像素一致。

Windows / macOS 共用本源码 ZIP，Python 3.12+。MacBook 只装 WPS 即可查看离线报告；安装后导入新生成的 `mcp.local.json`，MCP 启动参数保持 `["-m", "xmlmcp"]`。

- 六种图名、坐标轴和表头为中文，均值及 COV 从逐点数据独立计算，标准差分母为 n−1。
- 原图锚点、对象属性、批注、复选框、VML 等保留；未修改部件及修改片段外字节核验不变。
- 单点明细表的平均值和 COV 行固定，空余测量位置留空，不移动样式行。
- 报告数据缺失继续留空并列出待补字段；新模板需首次确认映射。
- 随包包含常规和粗体中文标题字体及 OFL 许可。业务模板、委托单、原始 ZIP 和真实报告均不公开。

安装步骤见 [INSTALL.md](https://github.com/Veiluna/XMLMCP/blob/master/INSTALL.md)，映射和母版制作见 [报告说明](https://github.com/Veiluna/XMLMCP/blob/master/docs/report-workflow.md)。下载 portable ZIP 并在目标电脑运行安装脚本，不复制其他电脑的虚拟环境。
