import re
from google.cloud import vision as v
from ocr_parser_google import get_client, _extract_items, _group_by_y

# (group, pos, name, range, unit)
LEFT_TEMPLATE = [
    ("注塑設定", "S11",   "保壓切換位置",   "±2",   "mm"),
    ("注塑設定", "T39",   "冷卻時間",       "±2",   "s"),
    ("注塑設定", "S24",   "注射劑量位置",   "±2",   "mm"),
    ("注塑設定", "S9011", "塑化終止位置",   "±2",   "mm"),
    ("注塑設定", "S199",  "射出位置1",      "±2",   "mm"),
    ("注塑設定", "S200",  "射出位置2",      "±2",   "mm"),
    ("注塑設定", "S201",  "射出位置3",      "±2",   "mm"),
    ("注塑設定", "S202",  "射出位置4",      "±2",   "mm"),
    ("注塑設定", "S203",  "射出位置5",      "±2",   "mm"),
    ("注塑設定", "S204",  "射出位置6",      "±2",   "mm"),
    ("注塑設定", "S205",  "射出位置7",      "±2",   "mm"),
    ("注塑設定", "V200",  "射出速度1",      "±5",   "mm/s"),
    ("注塑設定", "V201",  "射出速度2",      "±5",   "mm/s"),
    ("注塑設定", "V202",  "射出速度3",      "±5",   "mm/s"),
    ("注塑設定", "V203",  "射出速度4",      "±5",   "mm/s"),
    ("注塑設定", "V204",  "射出速度5",      "±5",   "mm/s"),
    ("注塑設定", "V205",  "射出速度7",      "±5",   "mm/s"),
    ("注塑設定", "P117",  "保壓壓力1",      "±50",  "bar"),
    ("注塑設定", "P118",  "保壓壓力2",      "±50",  "bar"),
    ("注塑設定", "P119",  "保壓壓力3",      "±50",  "bar"),
    ("注塑設定", "T117",  "保壓時間1",      "±0.5", "s"),
    ("注塑設定", "T118",  "保壓時間2",      "±0.5", "s"),
    ("注塑設定", "T119",  "保壓時間3",      "±0.5", "s"),
    ("注塑設定", "T120",  "總保壓時間",     "±0.5", "s"),
    ("注塑設定", "T2",    "射出時間",       "±0.5", "s"),
    ("注塑設定", "P21",   "背壓",           "±10",  "bar"),
    ("注塑設定", "N21",   "螺桿轉速",       "±10",  "1/min"),
    ("注塑設定", "P12",   "切換液壓系統壓力","±50", "bar"),
    ("開針",    "TS-1",  "閥針打開延遲",   "±0.5", "s"),
    ("開針",    "T6.1",  "閥針關閉延遲",   "±0.5", "s"),
    ("開針",    "S1",    "注口",           "±10",  "%"),
    ("脫模",    "F111",  "生產週期時間",   "±2",   "s"),
    ("脫模",    "T71",   "頂出停留時間",   "±2",   "s"),
    ("冰水",    "",      "溫度",           "±5",   "℃"),
    ("冰水",    "",      "壓力",           "±5",   "kg-cm²"),
]

RIGHT_TEMPLATE = [
    ("機器加熱設定", "77",  "進料段1區溫度",     "±10", "℃"),
    ("機器加熱設定", "76",  "進料段2區溫度",     "±10", "℃"),
    ("機器加熱設定", "75",  "加熱段1區溫度",     "±10", "℃"),
    ("機器加熱設定", "74",  "加熱段2區溫度",     "±10", "℃"),
    ("機器加熱設定", "73",  "加熱段3區溫度",     "±10", "℃"),
    ("機器加熱設定", "72",  "塑化壓縮段溫度",    "±10", "℃"),
    ("機器加熱設定", "71",  "計量段溫度",        "±10", "℃"),
    ("機器加熱設定", "70",  "轉注射單元1區溫度", "±10", "℃"),
    ("機器加熱設定", "69",  "轉注射單元2區溫度", "±10", "℃"),
    ("機器加熱設定", "68",  "轉注射單元3區溫度", "±10", "℃"),
    ("機器加熱設定", "67",  "轉注射單元4區溫度", "±10", "℃"),
    ("機器加熱設定", "59",  "注射單元1區溫度",   "±10", "℃"),
    ("機器加熱設定", "58",  "注射單元2區溫度",   "±10", "℃"),
    ("機器加熱設定", "57",  "注射單元3區溫度",   "±10", "℃"),
    ("機器加熱設定", "56",  "轉換器單元1區溫度", "±10", "℃"),
    ("機器加熱設定", "55",  "轉換器單元2區溫度", "±10", "℃"),
    ("機器加熱設定", "54",  "轉換器單元3區溫度", "±10", "℃"),
    ("機器加熱設定", "53",  "轉換器單元4區溫度", "±10", "℃"),
    ("機器加熱設定", "51",  "射出頭溫度",        "±10", "℃"),
    ("模具加熱設定", "401", "灌口1區溫度",       "±10", "℃"),
    ("模具加熱設定", "402", "灌口2區溫度",       "±10", "℃"),
    ("模具加熱設定", "403", "主澆道區溫度",      "±10", "℃"),
    ("模具加熱設定", "404", "澆道1區溫度",       "±10", "℃"),
    ("模具加熱設定", "405", "澆道2區溫度",       "±10", "℃"),
    ("模具加熱設定", "406", "澆道3區溫度",       "±10", "℃"),
    ("模具加熱設定", "407", "澆道4區溫度",       "±10", "℃"),
    ("模具加熱設定", "408", "澆道5區溫度",       "±10", "℃"),
    ("模具加熱設定", "409", "澆道6區溫度",       "±10", "℃"),
    ("模具加熱設定", "410", "澆道7區溫度",       "±10", "℃"),
    ("模具加熱設定", "411", "澆道8區溫度",       "±10", "℃"),
    ("模具加熱設定", "412", "澆道9區溫度",       "±10", "℃"),
    ("模具加熱設定", "413", "澆道10區溫度",      "±10", "℃"),
    ("模具加熱設定", "414", "澆道11區溫度",      "±10", "℃"),
    ("模具加熱設定", "415", "澆道12區溫度",      "±10", "℃"),
    ("乾燥機",      "",    "溫度",              "±10", "℃"),
    ("乾燥機",      "",    "料位計(桶外長度)",   "",    "cm"),
    ("料桶",        "",    "桶子+暫存桶噸位",    "",    "T"),
]

SKIP_KW = {"項目", "位置", "名稱", "設定值", "設定", "範圍", "單位",
           "Netstal", "EM09", "EM13", "製程管制標準"}

# Position code pattern
_POS_PAT = re.compile(
    r"^(S\d+|T\d+[\.\-]\d*|T\d+|V\d+|P\d+|N\d+|F\d+|TS[\-]\d+|T6[\.\-]\d+|\d{2,3})$",
    re.IGNORECASE
)


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
        left_val_x  = mid_x * 0.55
        right_val_x = mid_x + mid_x * 0.35
        data_items  = items
    left_items  = [i for i in data_items if i["x"] < mid_x]
    right_items = [i for i in data_items if i["x"] >= mid_x]
    left_params  = _match_params(left_items,  LEFT_TEMPLATE,  left_val_x)
    right_params = _match_params(right_items, RIGHT_TEMPLATE, right_val_x)
    return {**meta, "params": left_params + right_params}


def _empty():
    params = [{"group": g, "pos": p, "name": n, "value": "", "range": r, "unit": u}
              for g, p, n, r, u in LEFT_TEMPLATE + RIGHT_TEMPLATE]
    return {"機型": "", "日期": "", "克重": "", "瓶口": "",
            "模號": "", "穴數": "", "原料": "", "params": params}


def _extract_meta(items):
    meta = {"機型": "", "日期": "", "克重": "", "瓶口": "",
            "模號": "", "穴數": "", "原料": ""}
    rows = _group_by_y(items, tol=15)
    for group in rows[:8]:
        line = "".join(i["text"] for i in group)
        if not meta["機型"]:
            for mtype in ["EM09", "EM13"]:
                idx = line.find(mtype)
                if idx >= 0:
                    before = line[max(0, idx-4):idx]
                    if any(c in before for c in ["■", "√", "✓", "V"]):
                        meta["機型"] = mtype
                        break
        if not meta["日期"]:
            m = re.search(r"(\d{4})[.\-/年](\d{1,2})[.\-/月](\d{1,2})", line)
            if m:
                meta["日期"] = f"{m.group(1)}.{m.group(2)}.{m.group(3)}"
        for key, pat in [
            ("克重", r"克重[:：]?\s*([0-9.]+\s*[Gg]?)"),
            ("瓶口", r"瓶口[:：]?\s*([A-Za-z0-9]+)"),
            ("模號", r"模號\s*([A-Za-z0-9\-]+)"),
            ("穴數", r"(\d+)\s*穴"),
            ("原料", r"原料[:：]?\s*([A-Za-z0-9\-]+)"),
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
    pos_to_value  = {}   # position code → value
    name_to_value = {}   # name → value (fallback)
    known_pos   = {t[1] for t in template if t[1]}
    known_names = [t[2] for t in template]

    for row in rows:
        row_text = "".join(i["text"] for i in row)
        if any(kw in row_text for kw in SKIP_KW):
            continue
        sorted_items = sorted(row, key=lambda i: i["x"])

        # Value: item closest to val_col_x containing digit
        value, best_dist = "", float("inf")
        for item in sorted_items:
            dist = abs(item["x"] - val_col_x)
            if dist < tol_x and dist < best_dist and re.search(r"[\d>]", item["text"]):
                best_dist, value = dist, item["text"]
        if not value:
            continue

        left_texts = [i["text"] for i in sorted_items if i["x"] < val_col_x - tol_x * 0.4]

        # Try position code match first
        matched_pos = None
        for txt in left_texts:
            clean = txt.strip()
            if clean in known_pos:
                matched_pos = clean
                break
            if _POS_PAT.match(clean) and clean in known_pos:
                matched_pos = clean
                break
        if matched_pos:
            pos_to_value[matched_pos] = value
            continue

        # Fallback: fuzzy name match
        name = "".join(left_texts)
        matched = _best_match(name, known_names)
        if matched:
            name_to_value[matched] = value

    result = []
    for g, p, n, r, u in template:
        val = pos_to_value.get(p, "") if p else name_to_value.get(n, "")
        if not val:
            val = name_to_value.get(n, "")
        result.append({"group": g, "pos": p, "name": n, "value": val, "range": r, "unit": u})
    return result


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
