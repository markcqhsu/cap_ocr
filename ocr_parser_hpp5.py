import re
from google.cloud import vision as v
from ocr_parser_google import get_client, _extract_items, _group_by_y

LEFT_TEMPLATE = [
    ("注塑設定", "注射切換速率",   "±2",    "%"),
    ("注塑設定", "注料量",         "±2",    "%"),
    ("注塑設定", "注塑時間限制",   "±2",    "s"),
    ("注塑設定", "快速加注率",     "±2",    "%"),
    ("注塑設定", "快速加注體積",   "±2",    "%"),
    ("注塑設定", "注料筒壓力限額", "±2",    "%"),
    ("注塑設定", "冷卻時間",       "±2",    "s"),
    ("保壓",    "保壓壓力1",       "±5",    "%"),
    ("保壓",    "保壓壓力2",       "±5",    "%"),
    ("保壓",    "保壓壓力3",       "±5",    "%"),
    ("保壓",    "保壓時間1",       "±0.5",  "s"),
    ("保壓",    "保壓時間2",       "±0.5",  "s"),
    ("保壓",    "保壓時間3",       "±0.5",  "s"),
    ("保壓",    "保壓時間合計",    "±0.5",  "s"),
    ("保壓",    "初始保壓",        "±0.5",  "%"),
    ("保壓",    "最終保壓",        "±0.5",  "%"),
    ("保壓",    "前箝壓-延遲",     "±2",    "mm"),
    ("保壓",    "前箝壓-滯留時間", "±0.05", "s"),
    ("增塑",    "螺桿轉速",        "±10",   "rpm"),
    ("增塑",    "背壓",            "±100",  "psi"),
    ("增塑",    "擠塑機螺桿位置",  "±2",    "mm"),
    ("輸料",    "輸料速率",        "±2",    "mm"),
    ("輸料",    "輸料速度",        "±2",    "mm/s"),
    ("輸料",    "壓力限度",        "±100",  "psi"),
    ("輸料",    "充填壓力",        "±100",  "psi"),
    ("輸料",    "卸模時間",        "±10",   "%"),
    ("輸料",    "注口",            "±10",   "%"),
    ("控制",    "生產週期",        "±10",   "s"),
    ("控制",    "原料溫度",        "±10",   "℃"),
    ("控制",    "射出時間",        "±2",    "s"),
    ("控制",    "轉換位置",        "±2",    "mm"),
    ("控制",    "頂針前移",        "±2",    "mm"),
    ("控制",    "頂出滯留",        "±1",    "s"),
    ("控制",    "噴嘴",            "±5",    "t"),
    ("模具水",  "溫度",            "±5",    "℃"),
    ("模具水",  "壓力",            "±5",    "kg-cm²"),
    ("機",      "溫度",            "±10",   "℃"),
    ("機",      "料位計(桶外花度)","",      "cm"),
]

RIGHT_TEMPLATE = [
    ("機器加熱設定", "料管1",    "±10", "℃"),
    ("機器加熱設定", "料管2",    "±10", "℃"),
    ("機器加熱設定", "料管3",    "±10", "℃"),
    ("機器加熱設定", "料管4",    "±10", "℃"),
    ("機器加熱設定", "料管5",    "±10", "℃"),
    ("機器加熱設定", "料管6",    "±10", "℃"),
    ("機器加熱設定", "料管7",    "±10", "℃"),
    ("機器加熱設定", "機筒頭",   "±10", "℃"),
    ("機器加熱設定", "料筒伸出", "±10", "℃"),
    ("機器加熱設定", "注塑室1",  "±10", "℃"),
    ("機器加熱設定", "注塑室2",  "±10", "℃"),
    ("機器加熱設定", "注塑室3",  "±10", "℃"),
    ("機器加熱設定", "分配閥",   "±10", "℃"),
    ("機器加熱設定", "接頭",     "±10", "℃"),
    ("機器加熱設定", "SB",       "±10", "℃"),
    ("機器加熱設定", "XM1",      "±10", "℃"),
    ("機器加熱設定", "XM2",      "±10", "℃"),
    ("機器加熱設定", "XM3",      "±10", "℃"),
    ("機器加熱設定", "XM4",      "±10", "℃"),
    ("機器加熱設定", "XM5",      "±10", "℃"),
    ("機器加熱設定", "XM6",      "±10", "℃"),
    ("機器加熱設定", "MM11",     "±10", "℃"),
    ("機器加熱設定", "MM12",     "±10", "℃"),
    ("機器加熱設定", "MM13",     "±10", "℃"),
    ("機器加熱設定", "MM14",     "±10", "℃"),
    ("機器加熱設定", "MM15",     "±10", "℃"),
    ("機器加熱設定", "MM16",     "±10", "℃"),
    ("機器加熱設定", "MM17",     "±10", "℃"),
    ("機器加熱設定", "MM18",     "±10", "℃"),
    ("機器加熱設定", "MM19",     "±10", "℃"),
    ("機器加熱設定", "MM20",     "±10", "℃"),
    ("機器加熱設定", "MM21",     "±10", "℃"),
    ("機器加熱設定", "MM22",     "±10", "℃"),
    ("模具加熱設定", "TB8",      "±10", "℃"),
    ("模具加熱設定", "TB9",      "±10", "℃"),
    ("模具加熱設定", "TB10",     "±10", "℃"),
]

SKIP_KW = {"項目", "名稱", "設定值", "設定", "範圍", "單位",
           "Hpp5", "EM16", "製程管制標準", "返回目錄"}


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
        xs = [i["x"] for i in items]
        mid_x = (min(xs) + max(xs)) / 2
        left_val_x  = mid_x * 0.65
        right_val_x = mid_x + mid_x * 0.40
        data_items  = items
    left_items  = [i for i in data_items if i["x"] < mid_x]
    right_items = [i for i in data_items if i["x"] >= mid_x]
    left_params  = _match_params(left_items,  LEFT_TEMPLATE,  left_val_x)
    right_params = _match_params(right_items, RIGHT_TEMPLATE, right_val_x)
    return {**meta, "params": left_params + right_params}


def _empty():
    params = [{"group": g, "name": n, "value": "", "range": r, "unit": u}
              for g, n, r, u in LEFT_TEMPLATE + RIGHT_TEMPLATE]
    return {"日期": "", "克重": "", "瓶口": "", "模號": "", "穴數": "", "原料": "",
            "params": params}


def _extract_meta(items):
    meta = {"日期": "", "克重": "", "瓶口": "", "模號": "", "穴數": "", "原料": ""}
    rows = _group_by_y(items, tol=15)
    for group in rows[:8]:
        line = "".join(i["text"] for i in group)
        if not meta["日期"]:
            m = re.search(r"(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})", line)
            if m:
                meta["日期"] = f"{m.group(1)}.{m.group(2)}.{m.group(3)}"
        for key, pat in [
            ("克重", r"克重[:：]?\s*([0-9.]+\s*[Gg]?)"),
            ("瓶口", r"瓶口[:：]?\s*([A-Za-z0-9]+)"),
            ("模號", r"模號\s*([A-Za-z0-9\-]+)"),
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
                lx, rx = val_items[0]["x"], val_items[1]["x"]
                return {"header_y": header_y, "left_val_x": lx,
                        "right_val_x": rx, "mid_x": (lx + rx) / 2}
    return None


def _match_params(items, template, val_col_x, tol_x=90):
    rows = _group_by_y(items, tol=15)
    name_to_value = {}
    known_names = [t[1] for t in template]
    for row in rows:
        row_text = "".join(i["text"] for i in row)
        if any(kw in row_text for kw in SKIP_KW):
            continue
        sorted_items = sorted(row, key=lambda i: i["x"])
        value, best_dist = "", float("inf")
        for item in sorted_items:
            dist = abs(item["x"] - val_col_x)
            if dist < tol_x and dist < best_dist and re.search(r"[\d>]", item["text"]):
                best_dist, value = dist, item["text"]
        if not value:
            continue
        name = "".join(i["text"] for i in sorted_items if i["x"] < val_col_x - tol_x * 0.4)
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
