# -*- coding: utf-8 -*-
"""
案例 01｜多子公司批量并表
=============================================
场景：集团有 12 家分公司，每家每月交一张 Excel/CSV 报表上来。
      财务要做的就是把这十几张表拼成一张总表，再按科目汇总。

手工做：打开 12 个文件 → 复制 → 粘贴 → 再透视。至少半小时，还容易漏。
用脚本：3 秒出结果，并且自动告诉你哪家公司交少了、哪家科目名称写错了。

运行：
    python 批量并表.py
"""

import os
import glob

import pandas as pd

# ---------- 1. 找到所有子公司报表 ----------
当前目录 = os.path.dirname(os.path.abspath(__file__))
数据目录 = os.path.join(os.path.dirname(当前目录), "示例数据")
所有文件 = sorted(glob.glob(os.path.join(数据目录, "子公司报表_*.csv")))

print(f"一共找到 {len(所有文件)} 家公司的报表，开始合并...\n")

# ---------- 2. 逐个读进来，记下是谁家的 ----------
表格列表 = []
for 文件 in 所有文件:
    公司名 = os.path.basename(文件).replace("子公司报表_", "").replace(".csv", "")
    # 科目代码强制按文字读取，否则 1001 会被当成数字，前面的 0 也会丢
    表 = pd.read_csv(文件, encoding="utf-8-sig", dtype={"科目代码": str})
    表["公司"] = 公司名          # 补一列公司名，免得读进来分不清是谁的
    表格列表.append(表)

总表 = pd.concat(表格列表, ignore_index=True)

print(f"合并完成，总行数：{len(总表)} 行\n")

# ---------- 3. 脏数据清洗 ----------
# 3.1 科目名称前后可能多了空格，统一去掉
总表["科目名称"] = 总表["科目名称"].astype(str).str.strip()

# 3.2 金额列如果混进了文字，强制转成数字，转不出来的填 0
总表["本期金额"] = pd.to_numeric(总表["本期金额"], errors="coerce").fillna(0)

# 3.3 检查有没有空值
空值统计 = 总表.isnull().sum()
if 空值统计.sum() > 0:
    print("注意，以下列存在缺失值：")
    print(空值统计[空值统计 > 0])
else:
    print("数据完整性检查：没发现缺失值 ✓")

# ---------- 4. 检查哪家公司交少了科目 ----------
print("\n--- 各公司上报科目数 ---")
上报统计 = 总表.groupby("公司")["科目代码"].nunique().sort_values()
print(上报统计.to_string())

标准科目数 = 上报统计.max()
缺漏公司 = 上报统计[上报统计 < 标准科目数]
if len(缺漏公司) > 0:
    print(f"\n⚠️ 以下公司科目数不足（标准应为 {标准科目数} 个），需要退回补报：")
    for 公司, 数量 in 缺漏公司.items():
        已有 = set(总表[总表["公司"] == 公司]["科目代码"])
        全部 = set(总表["科目代码"])
        print(f"   {公司}：只报了 {数量} 个，缺 {sorted(全部 - 已有)}")

# ---------- 5. 按科目汇总，生成合并报表 ----------
print("\n--- 集团合并科目汇总表 ---")
汇总 = (
    总表.groupby(["科目代码", "科目名称"], as_index=False)["本期金额"]
    .agg(合计金额="sum", 涉及公司数="count")
    .sort_values("科目代码")
)
汇总["合计金额"] = 汇总["合计金额"].round(2)
print(汇总.to_string(index=False))

# ---------- 6. 算几个老板最关心的数 ----------
def 取金额(科目代码):
    """按科目代码取合计金额，取不到就返回 0"""
    结果 = 汇总[汇总["科目代码"] == 科目代码]["合计金额"]
    return float(结果.iloc[0]) if len(结果) > 0 else 0.0

收入 = 取金额("6001")
成本 = 取金额("6401")
应收 = 取金额("1122")
应付 = 取金额("2202")

print("\n--- 集团关键指标 ---")
print(f"主营业务收入合计：{收入:,.2f} 元")
print(f"主营业务成本合计：{成本:,.2f} 元")
print(f"毛利率：{(收入 - 成本) / 收入 * 100:.2f}%" if 收入 else "毛利率：无法计算")
print(f"应收账款余额：{应收:,.2f} 元")
print(f"应付账款余额：{应付:,.2f} 元")

# ---------- 7. 把结果存成 Excel ----------
输出目录 = os.path.join(当前目录, "输出")
os.makedirs(输出目录, exist_ok=True)
输出路径 = os.path.join(输出目录, "集团合并报表.xlsx")

with pd.ExcelWriter(输出路径, engine="openpyxl") as writer:
    汇总.to_excel(writer, sheet_name="科目汇总", index=False)
    总表.to_excel(writer, sheet_name="合并明细", index=False)
    上报统计.to_excel(writer, sheet_name="各公司上报情况")

print(f"\n结果已保存到：{输出路径}")
print("里面有 3 个工作表：科目汇总 / 合并明细 / 各公司上报情况")
