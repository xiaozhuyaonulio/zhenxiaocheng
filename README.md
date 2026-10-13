# zhenxiaocheng
amazon广告投流

## 最新：用户汇总表的完善版

- [26.9汇总表完善版](user_summary_review/26.9汇总表_完善版.xlsx)：18个页签，先看结论和重点词看板；保留源数据、比例公式、价格代表ASIN及完整广告范围。
- [完善说明与分析结论](user_summary_review/26.9汇总表_完善说明与分析结论.md)：原稿核对结果、口径修正、数据证据与待验证建议。
- [更新与复算说明](user_summary_review/README.md)：原稿保存在reference/26.9汇总表_原稿.xlsx；源数据更新后可运行脚本刷新。
- [逐字段计算公式](user_summary_review/字段计算公式说明.xlsx)：重点词看板25列和总表71列，区分数值计算、计数汇总、原始取数与规则标签。

已从原始数据重新计算2026年9月SQP与SP广告报告，独立分析商品承接、降本和增长机会。

- [给带教的汇报稿](analysis/给带教的汇报稿.md)：先读这份，适合口头汇报。
- [完整独立分析](analysis/9月独立分析报告.md)：商品承接、广告效率、场景机会与测试建议。
- [独立分析工作簿](analysis/9月独立分析.xlsx)：20个页签，含证据层级、原始行号、待复核候选与搜索机会矩阵。
- [文件与复算说明](analysis/README.md)：完整CSV、Word报告、图表和复算脚本。

当前数据不能可靠推算自然流量占比，也不能将广告原表的开始日期当作每日发生日期。分析未修改原始数据或任何广告投放设置。

所有指标仅以raw_data原始数据为来源。7月工作文件只提供组织思路参考，不作为9月指标来源。

## 在另一台电脑继续

登录有仓库权限的GitHub账号，克隆或下载本仓库，然后在Codex中打开项目根目录。报告、分析表、原始数据和复算脚本都在仓库中。

`gaoxiaocheng.zip`为完整项目包，包含用户汇总表原稿和完善版、最新独立分析、原始数据和7月参考文件。下载并解压也可以继续使用；旧压缩包保留在Git提交历史中。

若需要重新计算，在项目根目录执行：

```bash
python -m pip install -r requirements.txt
python analysis/analyze.py
python analysis/build_search_matrix.py
python analysis/build_deliverables.py
python user_summary_review/review_user_summary.py
```

7月原始参考文件位于`reference/梅森罐7月.xlsx`。
