# -*- coding: utf-8 -*-
"""
案例 02｜应收账款账龄分析
=============================================
场景：月底老板问一句"外面还有多少钱没收回来？欠了多久？"
      财务得按 0-30 天、31-60 天、61-90 天、90 天以上分档统计，
      还得单独把有争议的单据挑出来。

手工做：Excel 里用 IF 嵌套 + 透视表，公式一长串，下个月还得重来。
用脚本：一次写好，以后每个月换份数据直接跑。

运行：
    python 账龄分析.py
"""

import os
from datetime import datetime

import pandas as pd

当前目录 = os.path.dirname(os.path.abspath(__file__))
数据目录 = os.path.dirname(当前目录)
数据路径 = os.path.join(数据目录, "示例数据", "应收账款明细.csv")

# ---------- 1. 读数据 ----------
明细 = pd.read_csv(数据路径, encoding="utf-8-sig")

# 基准日：可以改成任意一天，比如 2026-08-31
基准日 = pd.to_datetime("2026-08-31")
明细["开票日期"] = pd.to_datetime(明细["开票日期"])

print(f"共读入 {len(明细)} 笔应收单据\n")

# ---------- 2. 计算未收金额与欠款天数 ----------
明细["未收金额"] = (明细["应收金额"] - 明细["已收金额"]).round(2)
明细["欠款天数"] = (基准日 - 明细["开票日期"]).dt.days

# 已经收完的单据不参与账龄分析
未清 = 明细[明细["未收金额"] > 0.01].copy()
print(f"其中尚未收完的：{len(未清)} 笔，合计 {未清['未收金额'].sum():,.2f} 元\n")


# ---------- 3. 分档：给每笔欠款贴一个账龄标签 ----------
def 判断账龄(天数):
    if 天数 <= 30:
        return "① 0-30天"
    elif 天数 <= 60:
        return "② 31-60天"
    elif 天数 <= 90:
        return "③ 61-90天"
    elif 天数 <= 180:
        return "④ 91-180天"
    else:
        return "⑤ 180天以上"


未清["账龄区间"] = 未清["欠款天数"].apply(判断账龄)

# ---------- 4. 汇总账龄表 ----------
账龄汇总 = (
    未清.groupby("账龄区间", as_index=False)
    .agg(单据笔数=("单据编号", "count"), 未收金额=("未收金额", "sum"))
    .sort_values("账龄区间")
)
账龄汇总["未收金额"] = 账龄汇总["未收金额"].round(2)
账龄汇总["占比"] = (账龄汇总["未收金额"] / 账龄汇总["未收金额"].sum() * 100).round(2).astype(str) + "%"

print("--- 账龄汇总表 ---")
print(账龄汇总.to_string(index=False))

# ---------- 5. 按客户排名：谁欠最多 ----------
客户汇总 = (
    未清.groupby("客户名称", as_index=False)
    .agg(单据笔数=("单据编号", "count"), 未收金额=("未收金额", "sum"), 最长欠款天数=("欠款天数", "max"))
    .sort_values("未收金额", ascending=False)
)
客户汇总["未收金额"] = 客户汇总["未收金额"].round(2)
客户汇总["占比"] = (客户汇总["未收金额"] / 客户汇总["未收金额"].sum() * 100).round(2)

print("\n--- 客户欠款排行（前 10）---")
print(客户汇总.head(10).to_string(index=False))

# ---------- 6. 把有争议的单据单独拎出来 ----------
争议 = 未清[未清["是否争议"] == "是"]
if len(争议) > 0:
    print(f"\n⚠️ 有 {len(争议)} 笔单据处于争议中，合计 {争议['未收金额'].sum():,.2f} 元，需要重点跟进：")
    print(争议[["单据编号", "客户名称", "欠款天数", "未收金额"]].to_string(index=False))
else:
    print("\n没有处于争议中的单据 ✓")

# ---------- 7. 画一张账龄分布柱状图（可选，需要装 matplotlib）----------
try:
    import matplotlib.pyplot as plt
    from matplotlib import font_manager

    # 让图表能显示中文（Windows 用微软雅黑，Mac 用苹方）
    for 字体 in ["Microsoft YaHei", "SimHei", "PingFang SC", "Arial Unicode MS"]:
        if 字体 in set(f.name for f in font_manager.fontManager.ttflist):
            plt.rcParams["font.sans-serif"] = [字体]
            break
    plt.rcParams["axes.unicode_minus"] = False

    图, 轴 = plt.subplots(figsize=(8, 4.5))
    颜色列表 = ["#4CAF50", "#8BC34A", "#FFC107", "#FF9800", "#F44336"]
    轴.bar(账龄汇总["账龄区间"], 账龄汇总["未收金额"], color=颜色列表[: len(账龄汇总)])
    轴.set_title("应收账款账龄分布（截至 2026-08-31）", fontsize=14)
    轴.set_ylabel("未收金额（元）")
    for i, 值 in enumerate(账龄汇总["未收金额"]):
        轴.text(i, 值, f"{值:,.0f}", ha="center", va="bottom", fontsize=9)
    plt.tight_layout()

    输出目录 = os.path.join(当前目录, "输出")
    os.makedirs(输出目录, exist_ok=True)
    图路径 = os.path.join(输出目录, "账龄分布图.png")
    plt.savefig(图路径, dpi=150)
    print(f"\n账龄分布图已保存：{图路径}")
except ImportError:
    print("\n（没装 matplotlib，跳过画图。想画图的话执行：python -m pip install matplotlib）")

# ---------- 8. 导出 Excel ----------
输出目录 = os.path.join(当前目录, "输出")
os.makedirs(输出目录, exist_ok=True)
输出路径 = os.path.join(输出目录, "账龄分析结果.xlsx")

with pd.ExcelWriter(输出路径, engine="openpyxl") as writer:
    账龄汇总.to_excel(writer, sheet_name="账龄汇总", index=False)
    客户汇总.to_excel(writer, sheet_name="客户排行", index=False)
    未清.sort_values("欠款天数", ascending=False).to_excel(writer, sheet_name="明细按天数排序", index=False)
    if len(争议) > 0:
        争议.to_excel(writer, sheet_name="争议单据", index=False)

print(f"分析结果已保存到：{输出路径}")
