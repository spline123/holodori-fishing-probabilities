from __future__ import annotations

import io
import json
import math
import os
import re
from collections import defaultdict
from collections.abc import Iterable
from dataclasses import dataclass

import matplotlib

matplotlib.use("Agg")  # 无窗口环境也要能出图，必须早于 pyplot 导入
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.ticker import MultipleLocator
from matplotlib.transforms import Bbox
from openpyxl.drawing.image import Image as XLImage


# ---- 配置 ----

# 抽鱼权重必需，映射与坐标缺失时降级
LOTTERY_JSON = "data/ParkFishingFishLottery.json"
PERMANENCE_JSON = "data/ParkPermanenceAction.json"  # spot → facility 映射
COORD_JSON = "data/env_3d_park_area-a03-area-01-noon_data.json"  # facility 坐标

OUTPUT_XLSX = "holodori-fishing.xlsx"
MAP_SHEET = "map"
IMAGE_SCALE = 0.6  # 地图插入 Excel 时的缩放

BAITS = tuple(f"park_fishing_bait-{i:03d}" for i in range(1, 9))
PERIOD_ORDER = {"noon": 0, "night": 1}

COLOR_SCALE = 1.8


@dataclass(frozen=True)
class MapStyle:
    figsize: tuple[float, float] = (6.8, 8.4)
    dpi: int = 150
    grid_color: str = "#E3E3DA"
    frame_color: str = "#A6AEA0"
    muted_color: str = "#9AA5B1"
    highlight_color: str = "#E8772E"
    ink_color: str = "#2F3437"
    faded_color: str = "#6B7075"
    point_size: float = 30.0
    label_size: float = 9.0
    flip_z: bool = False
    annotate_all: bool = False


MAP_STYLE = MapStyle()


# ---- 通用工具 ----

def load_json(path: str, required: bool = True) -> dict | list | None:
    """读取 UTF-8 JSON；required=False 时文件缺失只告警并返回 None。"""
    if not os.path.isfile(path):
        if required:
            raise FileNotFoundError(f"必需文件不存在：{path}")
        print(f"⚠️ 缺少可选文件 {path}，将使用降级策略")
        return None
    with open(path, "r", encoding="utf-8") as file:
        return json.load(file)


def short_id(resource_id: str) -> str:
    """去掉前缀：park_fishing_spot-01-01 → spot-01-01。"""
    return resource_id.removeprefix("park_fishing_")


_FACILITY_PATTERN = re.compile(r"^(?:perm|facility)-[^-]+-(.+?)(?:-fishing)?$")


def facility_name(permanence_id: str) -> str:
    """取出 facility 短名：perm-a03-fishingpond1-fishing → fishingpond1。"""
    if not permanence_id:
        return ""
    match = _FACILITY_PATTERN.match(permanence_id)
    return match.group(1) if match else permanence_id


_INVALID_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")


def make_sheet_name(label: str, period: str) -> str:
    """拼出合法的 sheet 名（≤31 字符）：{label}_{period}，非法字符换成下划线。"""
    suffix = f"_{period}"[:31]
    return f"{_INVALID_SHEET_CHARS.sub('_', label)[:31 - len(suffix)]}{suffix}"


def probability_style(probability: float | None) -> str:
    """按概率深浅返回单元格样式：无概率时灰底灰字，概率越高底色越深。"""
    if pd.isna(probability) or probability <= 0:
        return "background-color: #F2F2F2; color: #AAAAAA"
    intensity = min(probability * COLOR_SCALE, 1.0)
    r = int(232 + intensity * (27 - 232))
    g = int(245 + intensity * (94 - 245))
    b = int(233 + intensity * (32 - 233))
    text = "white" if intensity > 0.65 else "black"
    return f"background-color: #{r:02X}{g:02X}{b:02X}; color: {text}"


def mapped_facilities(spots: set[str], spot_facilities: dict[str, str] | None) -> set[str]:
    """返回本次用到的 facility 名。"""
    if not spot_facilities:
        return set()
    return {spot_facilities[spot] for spot in spots if spot in spot_facilities}


# ---- 地图 ----

# 标签候选位置，按偏好排序：(偏移, 水平对齐, 垂直对齐)
LABEL_OFFSETS = (
    ((7, 3), "left", "bottom"),
    ((7, -3), "left", "top"),
    ((-7, 3), "right", "bottom"),
    ((-7, -3), "right", "top"),
    ((0, 9), "center", "bottom"),
    ((0, -9), "center", "top"),
)


def tick_step(span: float) -> float:
    """取 1/2/5×10ⁿ 的刻度间隔，让两轴共用一把尺子。"""
    raw = span / 9 if span > 0 else 1
    magnitude = 10 ** math.floor(math.log10(raw))
    for factor in (1, 2, 5, 10):
        if factor * magnitude >= raw:
            return factor * magnitude
    return 10 * magnitude


def plot_map(
    positions: dict[str, tuple[float, float]],
    highlight: Iterable[str] = (),
) -> io.BytesIO:
    """把 facility 坐标画成散点图，highlight 中的点放大、加光晕并标注名字。"""
    highlight = set(highlight)
    figure, axes = plt.subplots(figsize=MAP_STYLE.figsize, dpi=MAP_STYLE.dpi)

    muted = [(x, z) for name, (x, z) in positions.items() if name not in highlight]
    if muted:
        axes.scatter(*zip(*muted), s=MAP_STYLE.point_size, c=MAP_STYLE.muted_color,
                     edgecolors="white", linewidths=0.8, alpha=0.85, zorder=2)

    lit = [(name, positions[name]) for name in positions if name in highlight]
    if lit:
        xs, zs = zip(*(xy for _, xy in lit))
        axes.scatter(xs, zs, s=MAP_STYLE.point_size * 3.2, c=MAP_STYLE.highlight_color,
                     alpha=0.16, edgecolors="none", zorder=3)
        axes.scatter(xs, zs, s=MAP_STYLE.point_size * 2.0, c=MAP_STYLE.highlight_color,
                     edgecolors="white", linewidths=1.4, zorder=4)

    if positions:
        xs, zs = zip(*positions.values())
        step = tick_step(max(max(xs) - min(xs), max(zs) - min(zs)))
        axes.xaxis.set_major_locator(MultipleLocator(step))
        axes.yaxis.set_major_locator(MultipleLocator(step))

    axes.set_title("Facility Map", loc="left", pad=14,
                   fontsize=12, fontweight="bold", color=MAP_STYLE.ink_color)
    axes.grid(True, color=MAP_STYLE.grid_color, linewidth=0.8)
    axes.set_axisbelow(True)
    axes.tick_params(colors=MAP_STYLE.faded_color, labelsize=9, length=0)
    for side in axes.spines:
        axes.spines[side].set_color(MAP_STYLE.frame_color)
        axes.spines[side].set_linewidth(1.2)
    axes.margins(0.07)
    axes.set_aspect("equal", adjustable="datalim")
    if MAP_STYLE.flip_z:
        axes.invert_yaxis()
    figure.tight_layout()

    labelled = list(positions) if MAP_STYLE.annotate_all else [name for name, _ in lit]
    if labelled:
        figure.canvas.draw()
        renderer = figure.canvas.get_renderer()
        points = {}
        for name, (x, z) in positions.items():
            px, py = axes.transData.transform((x, z))
            points[name] = Bbox.from_extents(px - 10, py - 10, px + 10, py + 10)

        placed = []
        for name in labelled:
            blockers = [box for other, box in points.items() if other != name] + placed
            for offset, align_h, align_v in LABEL_OFFSETS:
                text = axes.annotate(
                    name, positions[name], textcoords="offset points", xytext=offset,
                    ha=align_h, va=align_v, zorder=6,
                    fontsize=MAP_STYLE.label_size, color=MAP_STYLE.ink_color,
                    bbox={"boxstyle": "round,pad=0.18", "fc": "white",
                          "ec": "none", "alpha": 0.82},
                )
                box = text.get_window_extent(renderer).expanded(1.06, 1.3)
                inside = axes.bbox.contains(box.x0, box.y0) and axes.bbox.contains(box.x1, box.y1)
                if inside and not any(box.overlaps(blocker) for blocker in blockers):
                    break
                if (offset, align_h, align_v) != LABEL_OFFSETS[-1]:
                    text.remove()
            placed.append(box)

    buffer = io.BytesIO()
    figure.savefig(buffer, format="png", bbox_inches="tight", pad_inches=0.12)
    plt.close(figure)
    buffer.seek(0)
    return buffer


# ---- 数据载入 ----

_SPOT_PATTERN = re.compile(r"spot-(\d+)-(\d+)")


def sheet_order(spot: str, period: str) -> tuple[int, int, int]:
    """返回 sheet 顺序键：spot 组编号 → 时段 → spot 序号。"""
    group, number = _SPOT_PATTERN.search(spot).groups()
    return int(group), PERIOD_ORDER.get(period, len(PERIOD_ORDER)), int(number)


def load_lottery(path: str) -> tuple[pd.DataFrame, int]:
    """展开抽鱼权重：每行一条 (spot, 时段, bait, fish) 记录，按 sheet 顺序排好。"""
    raw = load_json(path, required=True)

    weights = defaultdict(list)
    for item in raw:
        data = item["data"]
        group = (
            data["parkFishingSpotId"],
            data["parkFishingBaitId"],
            data["parkFishingTimePeriodId"],
        )
        weights[group].append((data["parkFishingFishId"], data["hitRatio"]))

    records = []
    for (spot, bait, period), fishes in weights.items():
        total = sum(ratio for _, ratio in fishes) or 1  # 权重全为 0 时避免除零
        records.extend(
            {
                "spot": spot,
                "period": period.removeprefix("park_fishing_time_period-"),
                "bait": bait,
                "fish": short_id(fish),
                "prob": ratio / total,
            }
            for fish, ratio in fishes
        )

    # 记录顺序即 sheet 顺序，write_workbook 的 groupby(sort=False) 直接沿用
    records.sort(key=lambda record: sheet_order(record["spot"], record["period"]))
    return pd.DataFrame(records), len(raw)


def load_spot_facilities(path: str) -> dict[str, str] | None:
    """读取 spot → facility 映射；文件缺失返回 None。"""
    raw = load_json(path, required=False)
    if raw is None:
        return None
    facilities = {}
    for item in raw:
        data = item.get("data", {})
        spot = data.get("parkFishingSpotId")
        facility_id = data.get("parkPermanenceId") or item.get("park_permanence_id")
        if spot and facility_id:
            facilities[spot] = facility_name(facility_id)
    return facilities


def load_facility_positions(path: str) -> dict[str, tuple[float, float]] | None:
    """读取 facility 坐标 {名称: (x, z)}；文件缺失返回 None。"""
    raw = load_json(path, required=False)
    if raw is None:
        return None
    positions = {}
    for item in raw.get("_facilityData", []):
        name = facility_name(item.get("facilityId", ""))
        if not name:
            continue
        position = item.get("position")
        if not isinstance(position, dict) or "x" not in position or "z" not in position:
            print(f"⚠️ facility '{name}' 缺少 position/x/z，跳过：{item}")
            continue
        positions[name] = (float(position["x"]), float(position["z"]))
    return positions


# ---- 报表 ----

def report_facility_coverage(spots: set[str], spot_facilities: dict[str, str] | None) -> None:
    if spot_facilities is None:
        print("facility: 未加载，所有 sheet 名回退为 spot 短名")
        return
    covered = spots & set(spot_facilities)
    print(f"facility: {len(spot_facilities)} 条映射, "
          f"覆盖 {len(covered)}/{len(spots)} spot")
    uncovered = spots - covered
    if uncovered:
        print(f"⚠️ {len(uncovered)} 个 spot 无 facility，"
              f"sheet 名将回退：{sorted(uncovered)[:5]}")


def report_name_coverage(fishes: set[str], fish_names: dict[str, str]) -> None:
    covered = fishes & set(fish_names)
    print(f"names:    {len(fish_names)} 条映射, "
          f"覆盖 {len(covered)}/{len(fishes)} fish")
    uncovered = fishes - covered
    if uncovered:
        print(f"⚠️ {len(uncovered)} 条鱼无显示名，"
              f"将保留原名：{sorted(uncovered)[:5]}")


def report_facility_positions(
    spots: set[str],
    spot_facilities: dict[str, str] | None,
    facility_positions: dict[str, tuple[float, float]] | None,
) -> None:
    if facility_positions is None:
        print("coords:   未加载，跳过地图 sheet")
        return
    facilities = mapped_facilities(spots, spot_facilities)
    covered = facilities & set(facility_positions)
    print(f"coords:   {len(facility_positions)} 个 facility 坐标, "
          f"覆盖 {len(covered)}/{len(facilities)} 本次使用的 facility")
    missing = facilities - covered
    if missing:
        print(f"⚠️ {len(missing)} 个 facility 在坐标文件中缺失，"
              f"无法在图上高亮：{sorted(missing)[:5]}")


def make_table(rows: pd.DataFrame) -> pd.DataFrame:
    """把一个 (spot, 时段) 的记录透视为 fish × bait 概率表，bait 列用短名。"""
    table = rows.pivot_table(
        index="fish", columns="bait", values="prob",
        aggfunc="sum", fill_value=0.0,
    )
    table = table.reindex(columns=list(BAITS), fill_value=0.0).sort_index()
    table.columns = [short_id(bait) for bait in table.columns]
    table.index.name = "fish"
    return table


def write_workbook(
    output_xlsx: str,
    df: pd.DataFrame,
    spot_facilities: dict[str, str] | None,
    fish_names: dict[str, str] | None,
    facility_positions: dict[str, tuple[float, float]] | None,
) -> int:
    """写出地图 sheet 与全部概率表，返回概率表数量。"""
    sheet_count = 0

    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
        workbook = writer.book
        if "Sheet" in workbook.sheetnames:
            del workbook["Sheet"]

        if facility_positions:
            highlight = mapped_facilities(set(df["spot"]), spot_facilities)
            map_buffer = plot_map(facility_positions, highlight & set(facility_positions))
            image = XLImage(map_buffer)
            if IMAGE_SCALE != 1.0:
                image.width = int(image.width * IMAGE_SCALE)
                image.height = int(image.height * IMAGE_SCALE)
            worksheet = workbook.create_sheet(title=MAP_SHEET, index=0)
            worksheet.add_image(image, "A1")
            worksheet.sheet_view.showGridLines = False

        # sort=False：沿用记录顺序；默认按分组键字符串重排，night 会排到 noon 前面
        for (spot, period), rows in df.groupby(["spot", "period"], sort=False):
            unexpected_baits = set(rows["bait"]) - set(BAITS)
            if unexpected_baits:
                print(f"⚠️ spot={short_id(spot)}, period={period} "
                      f"出现未预期的 bait：{sorted(unexpected_baits)}")

            display_name = (spot_facilities or {}).get(spot) or short_id(spot)
            sheet = make_sheet_name(display_name, period)
            if sheet in workbook.sheetnames:
                raise ValueError(f"sheet 名冲突：{sheet}")

            table = make_table(rows)
            if fish_names:
                table.index = [fish_names.get(fish, fish) for fish in table.index]
                table.index.name = "fish"
            (
                table.style
                .map(probability_style)
                .format("{:.2%}")
                .set_caption(f"{display_name} · {period}")
                .to_excel(writer, sheet_name=sheet)
            )
            sheet_count += 1

    return sheet_count


def generate_report(output_xlsx: str, fish_names: dict[str, str] | None = None) -> None:
    """生成整份报表；fish_names 为可选的显示名映射，缺省时保留 fish-XXXX。"""
    df, record_count = load_lottery(LOTTERY_JSON)
    spots = set(df["spot"])
    print(f"lottery:  {record_count} 条记录, {len(spots)} 个 spot, "
          f"{df['fish'].nunique()} 种 fish")

    spot_facilities = load_spot_facilities(PERMANENCE_JSON)
    facility_positions = load_facility_positions(COORD_JSON)

    report_facility_coverage(spots, spot_facilities)
    if fish_names is not None:
        report_name_coverage(set(df["fish"]), fish_names)
    report_facility_positions(spots, spot_facilities, facility_positions)

    sheet_count = write_workbook(
        output_xlsx, df, spot_facilities, fish_names, facility_positions)
    map_count = 1 if facility_positions else 0
    print(f"共生成 {sheet_count} 张概率表 + {map_count} 张 map → {output_xlsx}")


if __name__ == "__main__":
    generate_report(OUTPUT_XLSX)
