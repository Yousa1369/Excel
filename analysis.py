# -*- coding: utf-8 -*-
"""
分析脚本：美化表格、计算门店月收入/货款比/币值/出货比、可视化、相关性分析与预警门店输出。
用法（示例）:
    python analysis.py "月报数据（模拟非真实）.xlsx"
"""
import sys, os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
from scipy.stats import pearsonr

plt.style.use('seaborn-whitegrid')
sns.set_palette('tab10')

# 配置
INPUT_XLSX = sys.argv[1] if len(sys.argv) > 1 else "月报数据（模拟非真实）.xlsx"
OUTPUT_DIR = "analysis_outputs"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 预警阈值（可调整）
WARNING_REVENUE_DROP_RATIO = 0.30
WARNING_MIN_MONTH_REVENUE = 5000
WARNING_HUOKUAN_RATIO_LOW = 0.6
WARNING_SHIPMENT_RATIO_DROP = 0.30

# 读取 Excel，自动选择行/列最多的 sheet 作为主表
xls = pd.ExcelFile(INPUT_XLSX, engine='openpyxl')
sheets = xls.sheet_names
candidates = []
for sh in sheets:
    df_tmp = pd.read_excel(xls, sheet_name=sh, engine='openpyxl')
    candidates.append((sh, df_tmp.shape[0]*df_tmp.shape[1]))
candidates_sorted = sorted(candidates, key=lambda x: x[1], reverse=True)
main_sheet = candidates_sorted[0][0]
print("选择主表：", main_sheet)
df = pd.read_excel(xls, sheet_name=main_sheet, engine='openpyxl')

# 字段候选映射
cols = [c.strip() for c in df.columns.astype(str)]
candidates_dict = {
    'store': ['门店', '店铺', '店名', 'store', 'shop'],
    'date': ['日期', '月份', '月', 'time', 'date'],
    'revenue': ['收入', '销售额', '营业额', 'revenue', 'sales', 'amount'],
    'payment_goods': ['货款', '货款金额', '货款收款', 'goods_payment'],
    'currency': ['币别','币种','currency'],
    'shipment_qty': ['出货', '出货量', '出货数量', 'shipment', 'qty', 'quantity'],
    'order_qty': ['订货量', '订单量', 'order_qty', 'orders'],
    'cost': ['成本', 'cost'],
    'receivable': ['应收', '应收款', 'receivable', 'arrears']
}
def find_col(keywords, cols):
    for k in keywords:
        for c in cols:
            if k.lower() in c.lower():
                return c
    return None

mapped = {}
for std, keywords in candidates_dict.items():
    found = find_col(keywords, cols)
    if found is not None:
        mapped[std] = found

print("自动映射字段：", mapped)

# 基本美化
df.dropna(axis=0, how='all', inplace=True)
df.dropna(axis=1, how='all', inplace=True)
df.columns = [c.strip() for c in df.columns.astype(str)]
if 'store' in mapped and 'date' in mapped:
    df = df.drop_duplicates(subset=[mapped['store'], mapped['date']], keep='last')
else:
    df = df.drop_duplicates(keep='last')

# 尝试把金额列转为数值
for col in df.columns:
    if any(k in col.lower() for k in ['金额','价格','price','amount','cost','收入','销售','货款']):
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(',','').str.replace('￥','').str.replace('$',''), errors='coerce')

rev_col = mapped.get('revenue')
pay_col = mapped.get('payment_goods')
cur_col = mapped.get('currency')
ship_col = mapped.get('shipment_qty')
date_col = mapped.get('date')
store_col = mapped.get('store')

# 日期/月份标准化
if date_col is None:
    for c in df.columns:
        if '月' in c or '月份' in c:
            date_col = c
            break
if date_col is not None:
    try:
        df[date_col] = pd.to_datetime(df[date_col])
    except:
        pass
if date_col is not None:
    df['month'] = pd.to_datetime(df[date_col]).dt.to_period('M').dt.to_timestamp()
else:
    df['month'] = pd.Timestamp.now().replace(day=1)

# 聚合
agg_inputs = {}
if rev_col: agg_inputs['revenue'] = pd.NamedAgg(column=rev_col, aggfunc='sum')
if pay_col: agg_inputs['payment_goods'] = pd.NamedAgg(column=pay_col, aggfunc='sum')
if ship_col: agg_inputs['shipment_qty'] = pd.NamedAgg(column=ship_col, aggfunc='sum')

group_fields = ['month'] + ([store_col] if store_col else [])

if len(agg_inputs) == 0:
    raise SystemExit("未能识别可用于汇总的金额/出货列，请检查表头。")

monthly_store = df.groupby(group_fields).agg(**agg_inputs).reset_index()
if 'revenue' in monthly_store.columns and 'payment_goods' in monthly_store.columns:
    monthly_store['huokuan_ratio'] = monthly_store['payment_goods'] / monthly_store['revenue']
else:
    monthly_store['huokuan_ratio'] = np.nan

if 'shipment_qty' in monthly_store.columns:
    if 'order_qty' in mapped and mapped['order_qty'] in df.columns:
        order_inputs = { 'order_qty': pd.NamedAgg(column=mapped['order_qty'], aggfunc='sum') }
        monthly_orders = df.groupby(group_fields).agg(**order_inputs).reset_index()
        monthly_store = monthly_store.merge(monthly_orders, on=group_fields, how='left')
        monthly_store['shipment_ratio'] = monthly_store['shipment_qty'] / monthly_store['order_qty']
    else:
        monthly_store['shipment_ratio'] = monthly_store['shipment_qty'] / (monthly_store['shipment_qty'].max() + 1)
else:
    monthly_store['shipment_ratio'] = np.nan

# 可视化与保存
os.makedirs(OUTPUT_DIR, exist_ok=True)
monthly_total = monthly_store.groupby('month')['revenue'].sum().reset_index()
plt.figure(figsize=(10,5))
import seaborn as sns
sns.barplot(data=monthly_total, x='month', y='revenue')
plt.xticks(rotation=45)
plt.title('每月总收入')
plt.tight_layout()
plt.savefig(os.path.join(OUTPUT_DIR, 'monthly_total_revenue.png'), dpi=150)
plt.close()

TOP_N = 8
if store_col:
    latest_month = monthly_store['month'].max()
    latest = monthly_store[monthly_store['month']==latest_month]
    top_stores = latest.sort_values('revenue', ascending=False).head(TOP_N)[store_col].tolist()
    top_trend = monthly_store[monthly_store[store_col].isin(top_stores)].pivot_table(index='month', columns=store_col, values='revenue', aggfunc='sum').fillna(0)
    top_trend.plot(kind='line', figsize=(12,6))
    plt.title(f'Top {TOP_N} 门店收入趋势')
    plt.ylabel('收入')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, 'top_stores_trend.png'), dpi=150)
    plt.close()

plt.figure(figsize=(8,5))
sns.histplot(monthly_store['huokuan_ratio'].dropna(), bins=20)
plt.title('货款比 分布')
plt.tight_layout()
plt.savefig(os.path.join(OUTPUT_DIR, 'huokuan_ratio_hist.png'), dpi=150)
plt.close()

if cur_col and cur_col in df.columns:
    cur_pivot = df.groupby(['month', cur_col])[rev_col].sum().unstack(fill_value=0)
    cur_pivot.plot(kind='bar', stacked=True, figsize=(12,6))
    plt.title('按币种堆叠的每月收入')
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, 'currency_monthly_stacked.png'), dpi=150)
    plt.close()

if 'shipment_ratio' in monthly_store.columns:
    plt.figure(figsize=(7,6))
    sns.scatterplot(data=monthly_store, x='shipment_ratio', y='revenue', hue=store_col if store_col else None, legend=False)
    plt.title('出货比 vs 收入 散点图')
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, 'shipment_vs_revenue_scatter.png'), dpi=150)
    plt.close()

num_cols = monthly_store.select_dtypes(include=[np.number]).columns.tolist()
if len(num_cols) > 0:
    corr = monthly_store[num_cols].corr(method='pearson')
    plt.figure(figsize=(8,6))
    sns.heatmap(corr, annot=True, fmt=".2f", cmap='RdBu', center=0)
    plt.title('数值字段皮尔森相关系数矩阵')
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, 'correlation_heatmap.png'), dpi=150)
    plt.close()

corr_results = []
if 'revenue' in monthly_store.columns:
    for col in ['shipment_ratio', 'huokuan_ratio']:
        if col in monthly_store.columns:
            valid = monthly_store[['revenue', col]].dropna()
            if len(valid) > 2:
                r, p = pearsonr(valid['revenue'], valid[col])
                corr_results.append({'pair': f'revenue vs {col}', 'pearson_r': r, 'p_value': p, 'n': len(valid)})
corr_df = pd.DataFrame(corr_results)
corr_df.to_csv(os.path.join(OUTPUT_DIR, 'correlation_results.csv'), index=False)

# 预警逻辑
warnings = []
if store_col:
    monthly_store_sorted = monthly_store.sort_values([store_col, 'month'])
    monthly_store_sorted['revenue_lag'] = monthly_store_sorted.groupby(store_col)['revenue'].shift(1)
    monthly_store_sorted['revenue_mom_drop'] = (monthly_store_sorted['revenue_lag'] - monthly_store_sorted['revenue']) / (monthly_store_sorted['revenue_lag'].replace(0, np.nan))
    monthly_store_sorted['shipment_lag'] = monthly_store_sorted.groupby(store_col)['shipment_ratio'].shift(1)
    monthly_store_sorted['shipment_mom_drop'] = (monthly_store_sorted['shipment_lag'] - monthly_store_sorted['shipment_ratio']) / (monthly_store_sorted['shipment_lag'].replace(0, np.nan))
    latest_m = monthly_store_sorted['month'].max()
    latest_df = monthly_store_sorted[monthly_store_sorted['month'] == latest_m]
    for _, row in latest_df.iterrows():
        s = row[store_col]
        recs = []
        if pd.notna(row.get('revenue_mom_drop')) and row['revenue_mom_drop'] >= WARNING_REVENUE_DROP_RATIO and row['revenue'] <= WARNING_MIN_MONTH_REVENUE:
            recs.append(f"收入环比下降 {row['revenue_mom_drop']:.0%}, 当月收入 {row['revenue']:.2f}")
        if pd.notna(row.get('huokuan_ratio')) and row.get('huokuan_ratio') < WARNING_HUOKUAN_RATIO_LOW:
            recs.append(f"货款比过低 {row.get('huokuan_ratio'):.2f}")
        if pd.notna(row.get('shipment_mom_drop')) and row.get('shipment_mom_drop') >= WARNING_SHIPMENT_RATIO_DROP:
            recs.append(f"出货比环比下降 {row.get('shipment_mom_drop'):.0%}")
        if recs:
            warnings.append({store_col: s, 'month': latest_m, 'reasons': "; ".join(recs), 'revenue': row['revenue'], 'huokuan_ratio': row.get('huokuan_ratio'), 'shipment_ratio': row.get('shipment_ratio')})
warnings_df = pd.DataFrame(warnings)
warnings_df.to_csv(os.path.join(OUTPUT_DIR, 'store_warnings.csv'), index=False)
monthly_store.to_csv(os.path.join(OUTPUT_DIR, 'monthly_store_summary.csv'), index=False)
print("分析完成，输出保存在", OUTPUT_DIR)
