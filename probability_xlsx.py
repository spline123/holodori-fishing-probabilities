# by DeepSeek-V4.1-Flash

import json
from collections import defaultdict
import pandas as pd

PERIODS = ["noon", "night"]
BAITS   = [f"park_fishing_bait-{i:03d}" for i in range(1, 9)]
PERIOD_ORDER = {p: i for i, p in enumerate(PERIODS)}   # noon=0, night=1


def prob_to_color(val):
    if val is None or pd.isna(val) or val <= 0:
        return "background-color: #F2F2F2; color: #AAAAAA"
    intensity = min(val * 1.8, 1.0)
    r = int(232 + intensity * (27  - 232))
    g = int(245 + intensity * (94  - 245))
    b = int(233 + intensity * (32  - 233))
    text = "white" if intensity > 0.65 else "black"
    return f"background-color: #{r:02X}{g:02X}{b:02X}; color: {text}"


with open("ParkFishingFishLottery.json", encoding="utf-8") as f:
    data = json.load(f)

# 1. 组内归一化
grouped = defaultdict(list)
for item in data:
    d = item["data"]
    grouped[(d["parkFishingSpotId"],
             d["parkFishingBaitId"],
             d["parkFishingTimePeriodId"])].append(
        (d["parkFishingFishId"], d["hitRatio"]))

records = []
for (spot, bait, period), fishes in grouped.items():
    total = sum(r for _, r in fishes)
    for fish, ratio in fishes:
        records.append({
            "spot":   spot,
            "period": period.replace("park_fishing_time_period-", ""),
            "bait":   bait,
            "fish":   fish.replace("park_fishing_", ""),
            "prob":   ratio / (total or 1),
        })

df = pd.DataFrame(records)

# 2. 构建自定义排序键
# spot 形如 "park_fishing_spot-01-01"
df["spot_a"] = df["spot"].str.extract(r"spot-(\d+)-")[0].astype(int)
df["spot_b"] = df["spot"].str.extract(r"spot-\d+-(\d+)")[0].astype(int)
df["period_order"] = df["period"].map(PERIOD_ORDER)

df = df.sort_values(
    ["spot_a", "period_order", "spot_b"]
).reset_index(drop=True)

# 3. 按排好的顺序遍历分组
with pd.ExcelWriter("holodori-fishing.xlsx", engine="openpyxl") as writer:
    idx = 0
    # sort=False 保证沿用 df 的排序
    for (spot, period), sub in df.groupby(["spot", "period"], sort=False):

        table = sub.pivot_table(
            index="fish", columns="bait",
            values="prob", aggfunc="sum", fill_value=0.0,
        )
        for b in BAITS:
            if b not in table.columns:
                table[b] = 0.0
        table = table[BAITS]
        table.columns = [c.replace("park_fishing_bait-", "bait-")
                         for c in table.columns]
        table = table.sort_index()

        styled = (
            table.style
                 .map(prob_to_color)
                 .format("{:.2%}")
                 .set_caption(f"{spot.replace('park_fishing_', '')} · {period}")
        )

        idx += 1
        short_spot = spot.replace("park_fishing_", "")
        sheet = f"{idx:02d}_{short_spot}_{period}"[:31]
        styled.to_excel(writer, sheet_name=sheet)

print(f"共生成 {idx} 张表 → holodori-fishing.xlsx")