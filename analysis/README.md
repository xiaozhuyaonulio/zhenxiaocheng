# 交付文件说明

先看：给带教的汇报稿.md / .docx（可用于口头汇报）。

完整报告：9月独立分析报告.md / .docx。

核对数据：9月独立分析.xlsx，20个页签，包含证据层级和搜索机会矩阵。所有结果写为数值，无需等待公式重算。

四张PNG图可独立插入汇报。完整细节保留在CSV中：广告_组合词汇总（51,736行）、广告_全部词汇总（43,846行）、两表_全部词对照（44,454行）等。

复算需Python与pandas、numpy、openpyxl、matplotlib、python-docx；本环境已安装。项目根目录依次执行：

```bash
python analysis/analyze.py
python analysis/build_search_matrix.py
python analysis/build_deliverables.py
```

原始文件位于source_material/raw_data，是原zip解包副本。source_manifest.json记录原始文件哈希。所有指标仅由raw_data中的CSV与Excel原表计算。原始文件未改动。

7月工作文件仅作为组织思路参考，未作为9月指标来源。搜索机会矩阵由build_search_matrix.py直接使用9月独立计算结果生成。

重要口径：缺失不等于0；SQP样本购买计数不是全店订单；不求自然流量差值；不把日期区间当日数据；广告组合不等于ASIN；目标ACOS尚未设定，行动阈值仅筛选待复核候选。
