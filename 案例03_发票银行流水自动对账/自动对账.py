# -*- coding: utf-8 -*-
"""
案例 03｜发票台账 vs 银行流水 自动对账
=============================================
场景：一边是开票台账（我们开了多少票），一边是银行流水（实际到账多少）。
      月底要找出三类差异：
        A. 开了票但没到账  → 催收
        B. 到账了但金额对不上（扣了手续费）→ 找银行回单
        C. 到账了但没开票  → 补开票 / 确认是不是预收款

手工做：VLOOKUP 一张张拉，眼睛看瞎，还容易漏。
用脚本：自动匹配，三类差异分别列出来，直接给老板看。

运行：
    python 自动对账.py
"""

import os

import pandas as pd

当前目录 = os.path.dirname(os.path.abspath(__file__))
数据目录 = os.path.join(os.path.dirname(当前目录), "示例数据")

# ---------- 1. 读两张表 ----------
发票 = pd.read_csv(os.path.join(数据目录, "发票台账.csv"), encoding="utf-8-sig")
流水 = pd.read_csv(os.path.join(数据目录, "银行流水.csv"), encoding="utf-8-sig")

print(f"发票台账：{len(发票)} 张，合计 {发票['价税合计'].sum():,.2f} 元")
print(f"银行流水：{len(流水)} 笔，合计 {流水['到账金额'].sum():,.2f} 元\n")

# 统一一下列名，方便后面合并
发票 = 发票.rename(columns={"发票号码": "发票号"})
流水 = 流水.rename(columns={"备注发票号": "发票号"})
发票["发票号"] = 发票["发票号"].astype(str).str.strip()
流水["发票号"] = 流水["发票号"].astype(str).str.strip()

# ---------- 2. 以发票为主表，左连接银行流水 ----------
对账 = 发票.merge(流水[["发票号", "到账金额", "到账日期"]], on="发票号", how="left")
对账["到账金额"] = 对账["到账金额"].fillna(0)
对账["差异"] = (对账["价税合计"] - 对账["到账金额"]).round(2)


def 判断状态(行):
    if 行["到账金额"] == 0:
        return "A. 已开票未到账"
    elif abs(行["差异"]) < 0.01:
        return "✓ 正常核销"
    else:
        return "B. 金额不符"


对账["对账状态"] = 对账.apply(判断状态, axis=1)

# ---------- 3. 分类输出 ----------
状态统计 = 对账["对账状态"].value_counts()
print("--- 对账结果概览 ---")
for 状态, 笔数 in 状态统计.items():
    金额 = 对账[对账["对账状态"] == 状态]["价税合计"].sum()
    print(f"  {状态}：{笔数} 张，涉及金额 {金额:,.2f} 元")

# A 类：开了票没到账
未到账 = 对账[对账["对账状态"] == "A. 已开票未到账"]
if len(未到账) > 0:
    print(f"\n--- A 类：已开票未到账（共 {len(未到账)} 张，{未到账['价税合计'].sum():,.2f} 元）---")
    print("  这批要安排催收：")
    print(未到账[["发票号", "开票日期", "客户名称", "价税合计"]].to_string(index=False))

# B 类：金额对不上
金额不符 = 对账[对账["对账状态"] == "B. 金额不符"]
if len(金额不符) > 0:
    print(f"\n--- B 类：到账金额与开票金额不符（共 {len(金额不符)} 张）---")
    print("  多半是银行扣了手续费，需要找回单核对：")
    print(金额不符[["发票号", "客户名称", "价税合计", "到账金额", "差异"]].to_string(index=False))

# C 类：银行有账，发票台账里找不到
已匹配发票号 = set(发票["发票号"])
无票流水 = 流水[(~流水["发票号"].isin(已匹配发票号)) | (流水["发票号"] == "")]
if len(无票流水) > 0:
    print(f"\n--- C 类：到账了但没发票（共 {len(无票流水)} 笔，{无票流水['到账金额'].sum():,.2f} 元）---")
    print("  可能是预收款或无票收入，需要确认是否补开票：")
    print(无票流水[["流水号", "到账日期", "付款方", "到账金额"]].to_string(index=False))

# ---------- 4. 核销率 ----------
正常笔数 = 状态统计.get("✓ 正常核销", 0)
核销率 = 正常笔数 / len(发票) * 100
print(f"\n本期发票核销率：{核销率:.1f}%")

# ---------- 5. 导出 Excel ----------
输出目录 = os.path.join(当前目录, "输出")
os.makedirs(输出目录, exist_ok=True)
输出路径 = os.path.join(输出目录, "对账结果.xlsx")

with pd.ExcelWriter(输出路径, engine="openpyxl") as writer:
    对账.to_excel(writer, sheet_name="全量对账明细", index=False)
    if len(未到账) > 0:
        未到账.to_excel(writer, sheet_name="A_已开票未到账", index=False)
    if len(金额不符) > 0:
        金额不符.to_excel(writer, sheet_name="B_金额不符", index=False)
    if len(无票流水) > 0:
        无票流水.to_excel(writer, sheet_name="C_到账无发票", index=False)

print(f"对账结果已保存到：{输出路径}")
