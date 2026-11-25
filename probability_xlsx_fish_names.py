from probability_xlsx import generate_report, load_json, short_id

LANGUAGE = "Chs"
LANGUAGE_JSON = f"data/LangParkFishingFish_{LANGUAGE}.json"
OUTPUT_XLSX = f"holodori-fishing-{LANGUAGE.lower()}.xlsx"


def fish_key(language_id: str) -> str:
    """把语言表 id 还原成 fish 键：la-park_fishing_fish-0001-name → fish-0001。"""
    key = language_id.removeprefix("la-").removesuffix("-name")
    return short_id(key)


def load_fish_names(path: str) -> dict[str, str] | None:
    """读取 fish 显示名；文件缺失返回 None。"""
    raw = load_json(path, required=False)
    if raw is None:
        return None
    names = {}
    for item in raw:
        data = item.get("data", {})
        item_id = data.get("id") or item.get("id", "")
        text = data.get("text")
        if text and item_id.endswith("-name"):
            names[fish_key(item_id)] = text
    return names


if __name__ == "__main__":
    fish_names = load_fish_names(LANGUAGE_JSON)
    if fish_names is None:
        print(f"names:    未加载（{LANGUAGE_JSON}），所有行名保持 fish-XXXX")
    generate_report(OUTPUT_XLSX, fish_names=fish_names)
