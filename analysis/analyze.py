"""Independent audit and analysis. Only reads raw CSV/XLSX; never loads supplied pickle.
Run from repository root: python analysis/analyze.py
"""
from pathlib import Path
import json, re, hashlib
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / 'source_material'
OUT = ROOT / 'analysis'
OUT.mkdir(exist_ok=True)
def norm(v):
    return re.sub(r'\s+', ' ', str(v).strip().lower())
def ratio(a, b):
    return a / b.replace(0, np.nan) if isinstance(b, pd.Series) else (a / b if b else np.nan)

files = sorted((SRC/'raw_data/产品9月份搜索查询绩效').glob('*.csv'))
raw = pd.concat([pd.read_csv(p, skiprows=1).assign(ASIN=p.stem, 来源文件=p.name, 原始行号=lambda d:d.index+3) for p in files], ignore_index=True)
sq = raw.rename(columns={'搜索查询':'原始搜索词','搜索查询得分':'查询排名','搜索查询量':'市场搜索量',
    '曝光：曝光总量':'市场曝光','曝光量：ASIN 计数':'样本曝光','点击量：总次数':'市场点击','点击量：ASIN 数量':'样本点击',
    '购物车添加：总数':'市场加购','加购：ASIN 数量':'样本加购','购买：下单总数':'市场购买','下单成交：ASIN 数量':'样本购买'})
sq['搜索词'] = sq['原始搜索词'].map(norm)
mp = pd.read_excel(SRC/'raw_data/对应关系.xlsx').dropna(subset=['ASIN'])
mp['ASIN'] = mp['ASIN'].str.strip()
assert not mp['ASIN'].duplicated().any()
sq = sq.merge(mp[['ASIN','品名','颜色','规格','型号']], on='ASIN', how='left', validate='many_to_one')
assert sq['品名'].notna().all()
assert not sq.duplicated(['ASIN','搜索词']).any()
mcols = ['市场搜索量','市场曝光','市场点击','市场加购','市场购买']
ocols = ['样本曝光','样本点击','样本加购','样本购买']
assert not sq.groupby('搜索词')[mcols].nunique().gt(1).any().any()
market = sq.drop_duplicates('搜索词')[['搜索词']+mcols].copy()
sq['点击对应市场预期购买'] = sq['样本点击'] * ratio(sq['市场购买'],sq['市场点击'])

adraw = pd.read_excel(SRC/'raw_data/26.9.xlsx')
ad = adraw.rename(columns={'客户搜索词':'原始客户搜索词','展示量':'广告曝光','点击量':'广告点击','花费':'广告花费',
    '7天总销售额':'广告销售额','7天总订单数(#)':'广告订单','7天总销售量(#)':'广告销量',
    '7天内广告SKU销售额':'广告SKU销售额','7天内其他SKU销售额':'其他SKU销售额'})
ad['原始行号'] = np.arange(2,len(ad)+2)
ad['搜索词'] = ad['原始客户搜索词'].map(norm)
ad['客户词类型'] = np.where(ad['搜索词'].str.fullmatch(r'b0[a-z0-9]{8}'), 'ASIN编码', '文本词')
auto = {'close-match':'自动-紧密','loose-match':'自动-宽泛','substitutes':'自动-同类','complements':'自动-关联'}
def target_type(r):
    if r['投放'] in auto: return auto[r['投放']]
    if r['匹配类型'] in ['EXACT','PHRASE','BROAD']:
        return {'EXACT':'手动-精准','PHRASE':'手动-词组','BROAD':'手动-广泛'}[r['匹配类型']]
    if str(r['投放']).startswith('asin'): return '商品定向-ASIN'
    if str(r['投放']).startswith('category'): return '商品定向-类目'
    if str(r['投放']).startswith('keyword-group'): return '关键词组定向'
    return '其他定向-需核查'
ad['投放类型'] = ad.apply(target_type,axis=1)
acols = ['广告曝光','广告点击','广告花费','广告销售额','广告订单','广告销量','广告SKU销售额','其他SKU销售额']
def rates(d):
    d=d.copy()
    d['CTR']=ratio(d['广告点击'],d['广告曝光'])
    d['CPC']=ratio(d['广告花费'],d['广告点击'])
    d['CVR']=ratio(d['广告订单'],d['广告点击'])
    d['ACOS']=ratio(d['广告花费'],d['广告销售额'])
    d['ROAS']=ratio(d['广告销售额'],d['广告花费'])
    d['广告SKU销售额口径ACOS']=ratio(d['广告花费'],d['广告SKU销售额'])
    d['其他SKU销售额占比']=ratio(d['其他SKU销售额'],d['广告销售额'])
    return d
def agg(keys):
    d=ad.groupby(keys,dropna=False)[acols].sum().reset_index()
    return rates(d).sort_values('广告花费',ascending=False).reset_index(drop=True)
portfolio=agg(['广告组合名称'])
portfolio['花费占比']=portfolio['广告花费']/ad['广告花费'].sum()
portfolio['销售额占比']=portfolio['广告销售额']/ad['广告销售额'].sum()
targettypes=agg(['投放类型'])
querytypes=agg(['客户词类型'])
adterms=agg(['搜索词','客户词类型'])
pq=agg(['广告组合名称','搜索词','客户词类型'])
targets=agg(['广告组合名称','广告活动名称','广告组名称','投放','匹配类型','投放类型'])
trace=ad.groupby(['广告组合名称','搜索词']).agg(原始行数=('原始行号','size'),
    首条原始行号=('原始行号','min'),末条原始行号=('原始行号','max'),
    开始日期=('开始日期','min'),结束日期=('结束日期','max')).reset_index()
pq=pq.merge(trace,on=['广告组合名称','搜索词'],validate='one_to_one')

asin=sq.groupby(['ASIN','品名','颜色','规格'],dropna=False).agg(
    收录词数=('搜索词','size'),样本曝光=('样本曝光','sum'),样本点击=('样本点击','sum'),
    样本加购=('样本加购','sum'),样本购买=('样本购买','sum'),市场曝光=('市场曝光','sum'),
    市场点击=('市场点击','sum'),市场购买=('市场购买','sum'),点击对应市场预期购买=('点击对应市场预期购买','sum')).reset_index()
asin['达到100条']=np.where(asin['收录词数'].eq(100),'是；仍不能量化遗漏','否；不代表全量')
asin['点击率']=ratio(asin['样本点击'],asin['样本曝光'])
asin['点击后购买率']=ratio(asin['样本购买'],asin['样本点击'])
asin['点击后加购率']=ratio(asin['样本加购'],asin['样本点击'])
asin['购买每加购比']=ratio(asin['样本购买'],asin['样本加购'])
asin['按自身点击词结构市场购买率']=ratio(asin['点击对应市场预期购买'],asin['样本点击'])
asin['结构调整购买指数']=ratio(asin['样本购买'],asin['点击对应市场预期购买'])
asin['样本购买贡献占比']=asin['样本购买']/asin['样本购买'].sum()
asin=asin.sort_values('样本购买',ascending=False).reset_index(drop=True)

qs=sq.groupby('搜索词').agg(样本曝光=('样本曝光','sum'),样本点击=('样本点击','sum'),
    样本加购=('样本加购','sum'),样本购买=('样本购买','sum'),收录ASIN数=('ASIN','nunique')).reset_index().merge(market,on='搜索词',validate='one_to_one')
qs['样本点击率']=ratio(qs['样本点击'],qs['样本曝光'])
qs['市场点击每曝光比']=ratio(qs['市场点击'],qs['市场曝光'])
qs['样本点击后购买率']=ratio(qs['样本购买'],qs['样本点击'])
qs['市场点击后购买率']=ratio(qs['市场购买'],qs['市场点击'])
qs['点击份额']=ratio(qs['样本点击'],qs['市场点击'])
qs['购买份额']=ratio(qs['样本购买'],qs['市场购买'])
qs['相对市场购买指数']=ratio(qs['样本点击后购买率'],qs['市场点击后购买率'])
qs['当前点击按市场率预期购买']=qs['样本点击']*qs['市场点击后购买率']
qs['距市场率购买差额']=qs['当前点击按市场率预期购买']-qs['样本购买']
qs['诊断']=np.select([
    qs['样本点击'].lt(50),
    qs['相对市场购买指数'].lt(.7),
    qs['相对市场购买指数'].ge(1.2)&qs['购买份额'].lt(.05)&qs['样本购买'].ge(10)],
    ['样本小，先积累','点击后购买偏弱，查相关性与承接','购买效率较强且份额低，测试增长'],default='持续观察')
qs=qs.sort_values('样本点击',ascending=False).reset_index(drop=True)

cross=qs.merge(adterms,on='搜索词',how='outer',indicator=True,validate='one_to_one')
cross['集合状态']=cross['_merge'].astype(str).map({'both':'两表都收录','left_only':'仅SQP收录','right_only':'仅广告收录'})
cross=cross.drop(columns='_merge')
cross['解释边界']='仅对照；不可相减求自然量，缺失不代表0'
coverage=cross.groupby('集合状态').agg(词条数=('搜索词','size'),样本点击=('样本点击','sum'),
    广告点击=('广告点击','sum'),广告花费=('广告花费','sum'),广告销售额=('广告销售额','sum'),广告订单=('广告订单','sum')).reset_index()
coverage['广告点击占全表比']=coverage['广告点击']/ad['广告点击'].sum()
coverage.loc[coverage['集合状态'].eq('仅广告收录'),'样本点击']=np.nan
coverage.loc[coverage['集合状态'].eq('仅SQP收录'),['广告点击','广告花费','广告销售额','广告订单','广告点击占全表比']]=np.nan
absent=cross[cross['集合状态'].eq('仅广告收录')].groupby('客户词类型')[acols].sum().reset_index()
absent['词条数']=cross[cross['集合状态'].eq('仅广告收录')].groupby('客户词类型').size().values

# Heuristic queues for review, not instructions to change advertising immediately.
pq['建议队列']=np.select([
    pq['广告订单'].eq(0)&pq['广告点击'].ge(20)&pq['广告花费'].ge(20),
    pq['广告订单'].ge(5)&pq['广告花费'].ge(50)&pq['ACOS'].gt(.4),
    pq['广告订单'].ge(10)&pq['ACOS'].le(.25)],
    ['优先复核：零单≥20点击且≥$20','复核：有单但ACOS>40%','增长测试：≥10单且ACOS≤25%'], default='观察')
queues=pq[pq['建议队列'].ne('观察')].copy()
aq=sq.copy()
aq['点击后购买率']=ratio(aq['样本购买'],aq['样本点击'])
aq['市场点击后购买率']=ratio(aq['市场购买'],aq['市场点击'])
aq['购买指数']=ratio(aq['点击后购买率'],aq['市场点击后购买率'])
aq['点击按市场率购买差额']=aq['点击对应市场预期购买']-aq['样本购买']

checks=[]
def check(name,value,notes): checks.append({'核对项目':name,'结果':value,'解释':notes})
check('分析数据来源','仅raw_data','读取51个SQP CSV、26.9.xlsx和对应关系.xlsx；不读取任何已有合并表、分析表或处理缓存。')
check('重复ASIN×归一化词',int(sq.duplicated(['ASIN','搜索词']).sum()),'无重复，ASIN计数可在样本内求和。')
check('同一词市场5项计数冲突',0,'所有重复出现词的市场计数一致，市场总量可按词去重。')
check('广告无日期业务键重复',int(ad.duplicated(['广告组合名称','广告活动名称','广告组名称','投放','匹配类型','原始客户搜索词']).sum()),'每个业务组合一行，支持区间汇总判断；不能当日表。')
check('广告开始日期不等于结束日期',int(ad['开始日期'].ne(ad['结束日期']).sum()),'这些行的区间花费不属于开始日期当日；不得均摊伪造日度数据。')
check('广告曝光0但点击>0',int((ad['广告曝光'].eq(0)&ad['广告点击'].gt(0)).sum()),'保留原始计数；该行CTR为空，不填0。需后台核查。')
pricecols=[c for c in raw if '价格' in c or '售价' in c]
pricebad=raw[raw[pricecols].gt(1000).any(axis=1)].copy()
check('SQP价格任一字段>1000的行数',len(pricebad),'玻璃容器价格中出现约6万—9万异常值；只标记，不自行修正，不用于价格策略。')
check('SQP份额四舍五入校验','通过','四项ASIN份额与计数比值的差均≤0.005个百分点；四舍五入不证明计数脱敏。')
for own,mc,share in [('样本曝光','市场曝光','曝光：ASIN 曝光占比 %'),('样本点击','市场点击','点击量：ASIN 份额 %'),('样本加购','市场加购','加入购物车：ASIN 加购占比 %'),('样本购买','市场购买','下单: ASIN 占比 %')]:
    err=(100*ratio(sq[own],sq[mc])-sq[share]).abs().dropna()
    assert err.max()<=.0050001,(share,err.max())
for col,count in [('点击量：点击率 %','市场点击'),('加入购物车：加购率%','市场加购'),('下单成交：成交转化率%','市场购买')]:
    err=(100*ratio(sq[count],sq['市场搜索量'])-sq[col]).abs().dropna()
    assert err.max()<=.0050001,(col,err.max())
assert np.allclose(ad['广告销售额'],ad['广告SKU销售额']+ad['其他SKU销售额'],atol=.011)
check('广告总销售额=广告SKU+其他SKU','逐行通过','总ACOS包含其他SKU销售额；不是广告商品自身盈利指标。')

tot=rates(pd.DataFrame([ad[acols].sum()])).iloc[0]
stats={
    'sqp_files':len(files),'sqp_rows':len(sq),'sqp_queries':sq['搜索词'].nunique(),'sqp_capped':int(sq.groupby('ASIN').size().eq(100).sum()),
    'sqp_totals':sq[ocols].sum().to_dict(),'market_unique_totals':market[mcols].sum().to_dict(),
    'ad_rows':len(ad),'ad_raw_queries':ad['原始客户搜索词'].nunique(),'ad_queries':ad['搜索词'].nunique(),
    'ad_totals':tot.to_dict(),'ad_date_range':[str(ad['开始日期'].min().date()),str(ad['结束日期'].max().date())],
    'ad_interval_rows':int(ad['开始日期'].ne(ad['结束日期']).sum()),
    'ad_interval_spend':float(ad.loc[ad['开始日期'].ne(ad['结束日期']),'广告花费'].sum()),
    'mapping_asins':len(mp),'missing_sqp_asins':sorted(set(mp['ASIN'])-set(sq['ASIN'])),
    'ad_zero_orders_spend':float(ad.loc[ad['广告订单'].eq(0),'广告花费'].sum()),
    'pq_zero_orders_spend':float(pq.loc[pq['广告订单'].eq(0),'广告花费'].sum()),
    'review_zero_count':int(pq['建议队列'].str.startswith('优先').sum()),
    'review_zero_spend':float(pq.loc[pq['建议队列'].str.startswith('优先'),'广告花费'].sum()),
    'price_bad_rows':len(pricebad),
    'coverage':coverage.to_dict(orient='records'),'absent_by_type':absent.to_dict(orient='records'),
}
def json_safe(v):
    if isinstance(v, dict): return {k:json_safe(x) for k,x in v.items()}
    if isinstance(v, list): return [json_safe(x) for x in v]
    if isinstance(v, (float,np.floating)) and not np.isfinite(v): return None
    return v
OUT.joinpath('metrics.json').write_text(json.dumps(json_safe(stats),ensure_ascii=False,indent=2,default=str,allow_nan=False))
frames={'SQP_ASIN汇总':asin,'SQP_搜索词诊断':qs,'SQP_ASIN词明细':aq,'广告_组合汇总':portfolio,
    '广告_投放类型':targettypes,'广告_客户词类型':querytypes,'广告_全部词汇总':adterms,
    '广告_组合词汇总':pq,'广告_投放单元汇总':targets,'两表_集合覆盖':coverage,
    '两表_全部词对照':cross,'行动_待复核队列':queues,'质量_核对记录':pd.DataFrame(checks),
    '质量_价格异常':pricebad,'范围_未导出ASIN':mp[~mp['ASIN'].isin(sq['ASIN'])]}
for name,d in frames.items(): d.to_csv(OUT/(name+'.csv'),index=False,encoding='utf-8-sig')

# Source fingerprint manifest makes a rerun auditable.
manifest=[]
for f in files+[SRC/'raw_data/26.9.xlsx',SRC/'raw_data/对应关系.xlsx']:
    manifest.append({'file':str(f.relative_to(ROOT)),'bytes':f.stat().st_size,'sha256':hashlib.sha256(f.read_bytes()).hexdigest()})
OUT.joinpath('source_manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2))
print(json.dumps(stats,ensure_ascii=False,indent=2,default=str))
print('\nPORTFOLIO\n',portfolio[['广告组合名称','广告花费','广告销售额','广告订单','ACOS','其他SKU销售额占比']].head(12).to_string(index=False))
print('\nTYPE\n',targettypes[['投放类型','广告花费','广告订单','ACOS','CVR','其他SKU销售额占比']].to_string(index=False))
print('\nASIN\n',asin[['ASIN','品名','样本点击','样本购买','点击后购买率','结构调整购买指数']].head(15).to_string(index=False))
print('\nZERO REVIEW\n',queues[queues['建议队列'].str.startswith('优先')][['广告组合名称','搜索词','广告点击','广告花费']].head(20).to_string(index=False))
print('\nHIGH ACOS\n',queues[queues['建议队列'].str.startswith('复核')][['广告组合名称','搜索词','广告点击','广告花费','广告订单','ACOS']].head(20).to_string(index=False))
print('\nGROWTH\n',queues[queues['建议队列'].str.startswith('增长')][['广告组合名称','搜索词','广告点击','广告花费','广告订单','ACOS']].head(20).to_string(index=False))
print('\nSQP GAPS\n',qs[qs['样本点击'].ge(100)].sort_values('距市场率购买差额',ascending=False)[['搜索词','样本点击','样本购买','市场点击后购买率','相对市场购买指数','距市场率购买差额']].head(20).to_string(index=False))
print('\nSQP GROWTH\n',qs[qs['诊断'].str.startswith('购买效率')][['搜索词','样本点击','样本购买','市场搜索量','购买份额','相对市场购买指数']].head(20).to_string(index=False))
