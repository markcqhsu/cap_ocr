import re
from google.cloud import vision as v
from ocr_parser_google import get_client, _extract_items, _group_by_y

# ── Fixed template (names / ranges / units are pre-printed) ───────────────────

LEFT_TEMPLATE = [
    ("注塑設定", "轉換位置",      "±2",   "mm"),
    ("注塑設定", "速度1",         "±5",   "mm"),
    ("注塑設定", "速度2",         "±5",   "mm"),
    ("注塑設定", "速度3",         "±5",   "mm"),
    ("注塑設定", "位置1(注塑量)", "±2",   "mm"),
    ("注塑設定", "位置2",         "±2",   "mm"),
    ("注塑設定", "位置3",         "±2",   "mm"),
    ("注塑",    "注入壓力限度",   "±200", "psi"),
    ("注塑",    "冷卻時間",       "±2",   "s"),
    ("注塑",    "射出時間",       "±2",   "s"),
    ("保壓",    "保壓壓力1",      "±50",  "psi"),
    ("保壓",    "保壓壓力2",      "±50",  "psi"),
    ("保壓",    "保壓壓力3",      "±50",  "psi"),
    ("保壓",    "保壓時間1",      "±0.5", "s"),
    ("保壓",    "保壓時間2",      "±0.5", "s"),
    ("保壓",    "保壓時間3",      "±0.5", "s"),
    ("保壓",    "保壓時間合計",   "±0.5", "s"),
    ("增塑",    "螺桿轉速",       "±10",  "rpm"),
    ("增塑",    "背壓",           "±100", "psi"),
    ("增塑",    "擠塑機向後位置", "±20",  "mm"),
    ("轉換",    "速度",           "±10",  "mm/s"),
    ("轉換",    "壓力限制",       "±100", "psi"),
    ("轉換",    "充填壓力",       "±100", "psi"),
    ("轉換",    "射料筒",         "±0.5", "s"),
    ("冰",      "溫度",           "±5",   "℃"),
    ("水",      "壓力",           "±5",   "kg-cm²"),
    ("分配器打開延遲", "分配器打開延遲", "±0.5", "s"),
]

RIGHT_TEMPLATE = [
    ("機器加熱設定", "擠塑機1",      "±10", "℃"),
    ("機器加熱設定", "擠塑機2",      "±10", "℃"),
    ("機器加熱設定", "擠塑機3",      "±10", "℃"),
    ("機器加熱設定", "擠塑機4",      "±10", "℃"),
    ("機器加熱設定", "擠塑機5",      "±10", "℃"),
    ("機器加熱設定", "擠塑機6",      "±10", "℃"),
    ("機器加熱設定", "擠塑機7",      "±10", "℃"),
    ("機器加熱設定", "擠塑機8",      "±10", "℃"),
    ("機器加熱設定", "B/尖-延伸",    "±10", "℃"),
    ("機器加熱設定", "注口接尖",     "±10", "℃"),
    ("機器加熱設定", "分配器",       "±10", "℃"),
    ("機器加熱設定", "S/鍋1",        "±10", "℃"),
    ("機器加熱設定", "S/鍋2",        "±10", "℃"),
    ("模具加熱設定", "S/B",          "±10", "℃"),
    ("模具加熱設定", "歧管1",        "±10", "℃"),
    ("模具加熱設定", "歧管2",        "±10", "℃"),
    ("模具加熱設定", "歧管3",        "±10", "℃"),
    ("模具加熱設定", "歧管4",        "±10", "℃"),
    ("模具加熱設定", "歧管5",        "±10", "℃"),
    ("模具加熱設定", "注口",         "±10", "%"),
    ("控制",        "頂出滯留",      "±2",  "s"),
    ("控制",        "生產週期",      "±2",  "s"),
    ("乾燥機",      "溫度",          "±10", "℃"),
    ("乾燥機",      "料位(西門子)",  "",    "%"),
    ("乾燥機",      "小料斗溫度",    "±10", "℃"),
    ("乾燥機",      "料桶噸位",      "",    "T"),
]

SKIP_KW = {"項目", "名稱", "設定值", "設定", "範圍", "單位", "Husky",
           "返回", "EM04", "EM06", "EM07", "製程管制標準"}


def parse_form(image_path: str) -> dict:
    with open(image_path, "rb") as f:
        content = f.read()

    image    = v.Image(content=content)
    response = get_client().document_text_detection(image=image)

    if response.error.message:
        raise RuntimeError(f"Vision API error: {response.error.message}")

    items = _extract_items(response)
    if not items:
        return _empty()

    items.sort(key=lambda i: (round(i["y"] / 10), i["x"]))

    meta   = _extract_meta(items)
    col_xs = _detect_columns(items)

    if col_xs:
        data_items  = [i for i in items if i["y"] > col_xs["header_y"] + 10]
        mid_x       = col_xs["mid_x"]
        left_val_x  = col_xs["left_val_x"]
        right_val_x = col_xs["right_val_x"]
    else:
        xs    = [i["x"] for i in items]
        mid_x = (min(xs) + max(xs)) / 2
        left_val_x  = mid_x * 0.65
        right_val_x = mid_x + mid_x * 0.40
        data_items  = items

    left_items  = [i for i in data_items if i["x"] < mid_x]
    right_items = [i for i in data_items if i["x"] >= mid_x]

    left_params  = _match_params(left_items,  LEFT_TEMPLATE,  left_val_x)
    right_params = _match_params(right_items, RIGHT_TEMPLATE, right_val_x)

    return {**meta, "params": left_params + right_params}


# ── helpers ────────────────────────────────────────────────────────────────────

def _empty():
    params = [{"group": g, "name": n, "value": "", "range": r, "unit": u}
              for g, n, r, u in LEFT_TEMPLATE + RIGHT_TEMPLATE]
    return {"機型": "", "日期": "", "克重": "", "瓶口": "", "模號": "", "穴數": "", "原料": "",
            "params": params}


def _extract_meta(items):
    meta = {"機型": "", "日期": "", "克重": "", "瓶口": "", "模號": "", "穴數": "", "原料": ""}
    rows = _group_by_y(items, tol=15)

    for group in rows[:8]:
        line = "".join(i["text"] for i in group)

        if not meta["機型"]:
            for mtype in ["EM04", "EM06", "EM07"]:
                idx = line.find(mtype)
                if idx >= 0:
                    before = line[max(0, idx - 4): idx]
                    if any(c in before for c in ["√", "✓", "V", "v", "☑"]):
                        meta["機型"] = mtype
                        break

        if not meta["日期"]:
            m = re.search(r"(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})", line)
            if m:
                meta["日期"] = f"{m.group(1)}.{m.group(2)}.{m.group(3)}"

        for key, pat in [
            ("克重", r"克重[:：]?\s*([0-9.]+\s*[Gg]?)"),
            ("瓶口", r"瓶口[:：]?\s*([A-Za-z0-9]+)"),
            ("模號", r"模號\s*([A-Za-z0-9\-\(\)]+)"),
            ("穴數", r"(\d+)\s*穴"),
            ("原料", r"原料[:：]?\s*([A-Za-z0-9]+)"),
        ]:
            if not meta[key]:
                m = re.search(pat, line)
                if m:
                    meta[key] = m.group(1).strip()

    return meta


def _detect_columns(items):
    for group in _group_by_y(items, tol=20):
        line = "".join(i["text"] for i in group)
        if "名稱" in line and "設定" in line and "範圍" in line:
            header_y  = group[0]["y"]
            val_items = sorted([i for i in group if "設定" in i["text"]], key=lambda i: i["x"])
            if len(val_items) >= 2:
                lx = val_items[0]["x"]
                rx = val_items[1]["x"]
                return {"header_y": header_y, "left_val_x": lx,
                        "right_val_x": rx, "mid_x": (lx + rx) / 2}
    return None


def _match_params(items, template, val_col_x, tol_x=90):
    """
    Group items by Y, extract (name, value) from each row,
    then fuzzy-match names to template and return ordered param list.
    """
    rows = _group_by_y(items, tol=15)
    name_to_value = {}
    known_names   = [t[1] for t in template]

    for row in rows:
        row_text = "".join(i["text"] for i in row)
        if any(kw in row_text for kw in SKIP_KW):
            continue

        sorted_items = sorted(row, key=lambda i: i["x"])

        # Value: item closest to val_col_x that contains a digit or >
        value, best_dist = "", float("inf")
        for item in sorted_items:
            dist = abs(item["x"] - val_col_x)
            if dist < tol_x and dist < best_dist and re.search(r"[\d>]", item["text"]):
                best_dist, value = dist, item["text"]

        if not value:
            continue

        # Name: concatenate items to the left of value column
        name = "".join(i["text"] for i in sorted_items
                       if i["x"] < val_col_x - tol_x * 0.4)

        matched = _best_match(name, known_names)
        if matched:
            name_to_value[matched] = value

    return [{"group": g, "name": n, "value": name_to_value.get(n, ""), "range": r, "unit": u}
            for g, n, r, u in template]


def _best_match(ocr_name, known_names, threshold=0.45):
    best, best_score = None, 0
    oc = set(ocr_name)
    for name in known_names:
        kn = set(name)
        overlap = len(oc & kn) / max(len(oc | kn), 1)
        if name in ocr_name or ocr_name in name:
            overlap += 0.25
        if overlap > best_score and overlap >= threshold:
            best_score, best = overlap, name
    return best
