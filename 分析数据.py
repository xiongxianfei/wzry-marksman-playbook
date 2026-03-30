"""
读取 训练记录.xlsx，生成分析图表。
运行：python 分析数据.py
输出：分析报告.png
"""

import sys
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.gridspec import GridSpec
import warnings
warnings.filterwarnings("ignore")

FILE   = "训练记录.xlsx"
OUTPUT = "分析报告.png"

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS"]
plt.rcParams["axes.unicode_minus"] = False

BG    = "#0D1117"
CARD  = "#161B22"
BORD  = "#21262D"
GOLD  = "#F0C040"
GREEN = "#3FB97A"
RED   = "#F85149"
GRAY  = "#8B949E"
WHITE = "#C9D1D9"

HERO_COLORS = {
    "后羿":   "#60A5FA",
    "莱西奥": "#A78BFA",
    "艾琳":   "#F472B6",
    "戈娅":   "#34D399",
    "孙尚香": "#FB923C",
}

YES_COLS = ["暴君", "三问", "靠队伍", "推塔", "跟强打", "心态"]


def calc_score(row):
    s  = 15 if pd.notna(row["经济"])  and row["经济"]  >= 6500 else 0
    s += 10 if pd.notna(row["死亡"])  and row["死亡"]  <= 2    else 0
    s += 10 if row["暴君"]  == "是" else 0
    s += 10 if row["三问"]  == "是" else 0
    s += 10 if row["靠队伍"] == "是" else 0
    s += 10 if row["推塔"]  == "是" else 0
    s += 10 if row["跟强打"] == "是" else 0
    s += 15 if row["心态"]  == "是" else 0
    s += 10 if pd.notna(row["备注"]) and str(row["备注"]).strip() != "" else 0
    return s


def load_data():
    try:
        df = pd.read_excel(FILE, sheet_name="每局记录", header=1, skiprows=[0])
    except FileNotFoundError:
        print(f"找不到 {FILE}，请先运行 生成模板.py 并填写数据。")
        sys.exit(1)

    df.columns = ["日期","英雄","胜负","经济","死亡","暴君","三问","靠队伍","推塔","跟强打","心态","_公式分","备注"]
    df = df[df["英雄"].notna() & (df["英雄"] != "英雄")].copy()
    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    df["经济"] = pd.to_numeric(df["经济"], errors="coerce")
    df["死亡"] = pd.to_numeric(df["死亡"], errors="coerce")
    for c in YES_COLS:
        df[c] = df[c].astype(str).str.strip()
    df["复盘分"] = df.apply(calc_score, axis=1).astype(float)
    df = df.dropna(subset=["日期","英雄"]).reset_index(drop=True)
    return df


# ── 绘图辅助 ─────────────────────────────────────────
def style_ax(ax):
    ax.set_facecolor(CARD)
    for sp in ax.spines.values():
        sp.set_color(BORD)
    ax.tick_params(colors=GRAY, labelsize=8)
    ax.title.set_color(WHITE)
    ax.title.set_fontsize(11)


def plot_trend(ax, series, title, target, color):
    style_ax(ax)
    vals = series.dropna().values[-25:]
    xs   = np.arange(len(vals))
    ax.plot(xs, vals, color=color, lw=2, zorder=3)
    ax.fill_between(xs, vals, alpha=0.12, color=color)
    ax.axhline(target, color=GRAY, lw=1, ls="--", alpha=0.6)
    ax.text(len(vals)*0.98, target, f"目标{target}", color=GRAY,
            fontsize=8, va="bottom", ha="right")
    avg = vals.mean()
    ax.axhline(avg, color=GOLD, lw=1, ls=":", alpha=0.8)
    ax.text(0, avg, f" 均{avg:.0f}", color=GOLD, fontsize=8, va="top")
    colors = [GREEN if v >= target else RED for v in vals]
    ax.scatter(xs, vals, c=colors, s=35, zorder=4)
    ax.set_title(title, pad=8)
    ax.set_xlim(-0.5, max(len(vals)-0.5, 1))


def plot_hero_bar(ax, df):
    style_ax(ax)
    vc = df["英雄"].value_counts()
    if vc.empty:
        ax.text(0.5, 0.5, "暂无数据", color=GRAY, ha="center", va="center", transform=ax.transAxes)
        return
    names  = vc.index.tolist()
    counts = vc.values
    colors = [HERO_COLORS.get(h, GOLD) for h in names]
    bars = ax.bar(names, counts, color=colors, width=0.55, alpha=0.85)
    for bar, name in zip(bars, names):
        sub  = df[df["英雄"] == name]
        rate = (sub["胜负"] == "胜").sum() / len(sub) * 100 if len(sub) else 0
        ax.text(bar.get_x() + bar.get_width()/2,
                bar.get_height() + 0.08,
                f"{rate:.0f}%胜", ha="center", va="bottom", color=WHITE, fontsize=8)
    ax.set_title("各英雄场次与胜率", pad=8)


def plot_pie(ax, df):
    style_ax(ax)
    ax.axis("off")
    scores = df["复盘分"].dropna()
    if scores.empty:
        ax.text(0.5, 0.5, "暂无数据", color=GRAY, ha="center", va="center", transform=ax.transAxes)
        return
    good = (scores >= 80).sum()
    mid  = ((scores >= 60) & (scores < 80)).sum()
    bad  = (scores < 60).sum()
    vals   = [bad, mid, good]
    labels = [f"重点复盘\n{bad}局", f"正常\n{mid}局", f"优秀\n{good}局"]
    colors = [RED, GOLD, GREEN]
    wedges, texts, autos = ax.pie(
        vals, labels=labels, colors=colors, autopct="%1.0f%%",
        startangle=90, textprops={"color": WHITE, "fontsize": 8},
        wedgeprops={"edgecolor": BG, "linewidth": 2}
    )
    for at in autos:
        at.set_color(BG); at.set_fontweight("bold")
    ax.set_title("复盘分分布", pad=8, color=WHITE)


def plot_radar(ax, df):
    dims = ["经济达标", "死亡控制", "暴君参与", "决策三问", "推塔意识", "跟强打"]
    n = len(dims)

    def pct(col, threshold=None, target_val="是"):
        sub = df[col].dropna()
        if len(sub) == 0: return 0
        if threshold is not None:
            num = pd.to_numeric(df[col], errors="coerce").dropna()
            return (num >= threshold).sum() / len(num) * 100 if len(num) else 0
        return (sub == target_val).sum() / len(sub) * 100

    values = [
        pct("经济", threshold=6500),
        100 - pct("死亡", threshold=3),
        pct("暴君"),
        pct("三问"),
        pct("推塔"),
        pct("跟强打"),
    ]

    angles = np.linspace(0, 2*np.pi, n, endpoint=False).tolist()
    angles += angles[:1]
    values += values[:1]

    ax.set_facecolor(CARD)
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    ax.set_rlim(0, 100)
    ax.set_rticks([25, 50, 75, 100])
    ax.tick_params(colors=GRAY, labelsize=7)
    ax.grid(color=BORD, linewidth=0.8)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(dims, color=WHITE, fontsize=9)
    ax.spines["polar"].set_color(BORD)

    ax.plot(angles, values, color=GOLD, lw=2)
    ax.fill(angles, values, color=GOLD, alpha=0.2)
    ax.set_title("能力雷达图（%达标率）", color=WHITE, fontsize=11, pad=18)


def plot_stat_cards(ax, df):
    ax.set_facecolor(BG)
    ax.axis("off")
    n   = len(df)
    if n == 0: return
    wr  = (df["胜负"] == "胜").sum() / n * 100
    ec  = df["经济"].mean()
    d   = df["死亡"].mean()
    s   = df["复盘分"].mean()
    eco_ok = (df["经济"] >= 6500).sum() / n * 100

    items = [
        ("总场次",      f"{n} 局",      WHITE),
        ("胜率",        f"{wr:.1f}%",   GREEN if wr >= 50 else RED),
        ("均10分经济",  f"{ec:.0f}",    GREEN if ec >= 6500 else RED),
        ("均死亡",      f"{d:.1f} 次",  GREEN if d <= 2 else RED),
        ("均复盘分",    f"{s:.1f}",     GREEN if s >= 70 else (GOLD if s >= 60 else RED)),
        ("经济达标率",  f"{eco_ok:.1f}%", GREEN if eco_ok >= 70 else GOLD),
    ]

    for i, (label, val, color) in enumerate(items):
        x = (i % 3) / 3 + 0.04
        y = 0.72 if i < 3 else 0.18
        ax.text(x, y + 0.18, label, color=GRAY,  fontsize=10, transform=ax.transAxes)
        ax.text(x, y,        val,   color=color,  fontsize=22,
                fontweight="bold", transform=ax.transAxes)

    # 分隔线
    ax.axhline(0.5, color=BORD, lw=1, xmin=0.01, xmax=0.99)


# ── 主程序 ───────────────────────────────────────────
def main():
    df = load_data()
    if df.empty:
        print("暂无数据，请先填写训练记录。")
        sys.exit(0)
    print(f"读取到 {len(df)} 条记录，生成报告中...")

    fig = plt.figure(figsize=(16, 10), facecolor=BG)
    fig.suptitle("王者荣耀发育路训练分析报告",
                 color=GOLD, fontsize=15, fontweight="bold", y=0.98)

    # 手动指定位置避免极坐标冲突
    # 第一行：关键数字（整行）
    ax_stat = fig.add_axes([0.04, 0.76, 0.92, 0.18], facecolor=CARD)
    plot_stat_cards(ax_stat, df)
    for sp in ax_stat.spines.values(): sp.set_color(BORD)

    # 第二行：三个趋势图
    ax_ec  = fig.add_axes([0.04, 0.44, 0.28, 0.27])
    ax_sc  = fig.add_axes([0.37, 0.44, 0.28, 0.27])
    ax_d   = fig.add_axes([0.70, 0.44, 0.28, 0.27])
    plot_trend(ax_ec, df["经济"],   "10分钟经济趋势（近25局）", 6500, GOLD)
    plot_trend(ax_sc, df["复盘分"], "复盘分趋势（近25局）",     60,   GREEN)
    plot_trend(ax_d,  df["死亡"],   "死亡次数趋势（近25局）",   2,    RED)

    # 第三行：英雄柱状图 + 饼图 + 雷达
    ax_hero = fig.add_axes([0.04, 0.05, 0.28, 0.30])
    ax_pie  = fig.add_axes([0.37, 0.05, 0.28, 0.30])
    ax_rad  = fig.add_axes([0.67, 0.04, 0.32, 0.32], polar=True)

    plot_hero_bar(ax_hero, df)
    plot_pie(ax_pie, df)
    plot_radar(ax_rad, df)

    plt.savefig(OUTPUT, dpi=150, bbox_inches="tight", facecolor=BG)
    print(f"[OK] 报告已保存：{OUTPUT}")


if __name__ == "__main__":
    main()
