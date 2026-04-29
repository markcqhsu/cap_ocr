import re
import google.auth
from google.cloud import vision

_client = None

def get_client():
    global _client
    if _client is None:
        credentials, _ = google.auth.default(
            scopes=["https://www.googleapis.com/auth/cloud-vision"]
        )
        _client = vision.ImageAnnotatorClient(credentials=credentials)
    return _client


def parse_form(image_path: str) -> dict:
    with open(image_path, "rb") as f:
        content = f.read()

    image    = vision.Image(content=content)
    response = get_client().document_text_detection(image=image)

    if response.error.message:
        raise RuntimeError(f"Vision API error: {response.error.message}")

    items = _extract_items(response)
    if not items:
        return _empty()

    items.sort(key=lambda i: (round(i["y"] / 10), i["x"]))

    meta        = _extract_meta(items)
    col_xs, h_y = _detect_columns(items)
    footer_y    = _detect_footer_y(items)
    data_items  = [i for i in items if i["y"] > h_y + 30 and i["y"] < footer_y - 5]
    rows        = _extract_rows(data_items, col_xs)
    footer      = _extract_footer(items)

    return {**meta, "rows": rows, **footer}


def _empty():
    return {"製程": "", "組別": "", "班別": "", "年份": "", "rows": [], "廠長": "", "副廠": "", "填表人": ""}


def _extract_items(response):
    """word 層級擷取，每個 word 有獨立座標"""
    items = []
    for page in response.full_text_annotation.pages:
        for block in page.blocks:
            for para in block.paragraphs:
                for word in para.words:
                    text = "".join(s.text for s in word.symbols).strip()
                    if not text:
                        continue
                    verts = word.bounding_box.vertices
                    xs = [v.x for v in verts]
                    ys = [v.y for v in verts]
                    items.append({
                        "text": text,
                        "x": (min(xs) + max(xs)) / 2,
                        "y": (min(ys) + max(ys)) / 2,
                    })
    return items


# ── meta ──────────────────────────────────────────────────────────────────────

PROCESS_TYPES = ["切割", "射出", "組裝", "包裝", "印刷"]

def _extract_meta(items):
    meta = {"製程": "", "組別": "", "班別": "", "年份": ""}

    for p in PROCESS_TYPES:
        if any(p in i["text"] for i in items):
            meta["製程"] = p
            break

    # 同一行合併後再做 regex（Google Vision 會把「班別:C」拆成多個 word）
    rows = _group_by_y(items, tol=12)
    for group in rows[:10]:          # 只看表格上方的 meta 區域
        line = "".join(i["text"] for i in group)
        if "不良率紀錄" in line:     # 跳過標題行
            continue

        m = re.search(r"組別:?\s*(\d+)", line)
        if m and not meta["組別"]:
            meta["組別"] = m.group(1)

        m = re.search(r"班別:?\s*([A-Za-z])", line)
        if m and not meta["班別"]:
            meta["班別"] = m.group(1).upper()

        m = re.search(r"年份:?\s*(20\d{2})", line)
        if m:
            meta["年份"] = m.group(1)

    return meta


# ── column detection ───────────────────────────────────────────────────────────

# 每個欄位的辨識關鍵字（partial match）
def _detect_columns(items):
    """
    用「位置順序」識別欄位，處理 Google Vision 把相同文字拆散的問題：
    - '不良' 出現三次（不良數 / 不良率 / 不良原因），取出後按 X 排序
    - '廢' 出現多次，靠 '料' 旁邊的是廢料KG，靠 '蓋' 旁邊的是廢蓋KG
    """
    rows_by_y = _group_by_y(items, tol=30)   # header 跨 y=170~190，需較大 tol
    header_y, header_items = 0, []

    for group in rows_by_y:
        line = "".join(i["text"] for i in group)
        if sum(1 for kw in ["日期", "品名", "產量", "工單", "機台"] if kw in line) >= 3:
            header_y = group[0]["y"]
            header_items = group
            break

    if not header_items:
        return {}, 0

    col_xs = {}

    def first_x(pred):
        for i in sorted(header_items, key=lambda i: i["x"]):
            if pred(i["text"]):
                return i["x"]
        return None

    def nth_x(fragment, n):
        hits = [i for i in sorted(header_items, key=lambda i: i["x"])
                if fragment in i["text"]]
        return hits[n]["x"] if n < len(hits) else None

    col_xs["日期"]           = first_x(lambda t: "日期" in t)
    col_xs["機台編號"]       = first_x(lambda t: t in ("機", "機台") or "機台" in t)
    col_xs["工單號碼"]       = first_x(lambda t: t in ("工", "工單") or "工單" in t)
    col_xs["品名蓋型"]       = first_x(lambda t: "品名" in t)
    col_xs["產量"]           = first_x(lambda t: "產量" in t or t == "產")
    col_xs["不良數"]         = nth_x("不良", 0)   # 第一個「不良」
    col_xs["不良率"]         = nth_x("不良", 1)   # 第二個「不良」
    col_xs["有無添加回收料"] = first_x(lambda t: "有無" in t or "回收料" in t)
    col_xs["廢蓋KG"]         = nth_x("廢", 0)     # 第一個「廢」
    col_xs["廢料KG"]         = first_x(lambda t: "廢料" in t) or nth_x("廢", 1)
    col_xs["可回收廢蓋KG"]   = first_x(lambda t: "可回收" in t)   # 只匹配含「可回收」的 item
    col_xs["不良原因"]       = first_x(lambda t: "原因" in t) or nth_x("不良", 2)

    # 移除 None 的欄位
    col_xs = {k: v for k, v in col_xs.items() if v is not None}
    return col_xs, header_y


def _detect_footer_y(items):
    for item in items:
        if any(kw in item["text"] for kw in ["廠長", "副廠", "填表人"]):
            return item["y"]
    return float("inf")


# ── row extraction ─────────────────────────────────────────────────────────────

ALL_COLS = [
    "日期", "機台編號", "工單號碼", "品名蓋型", "產量", "不良數",
    "不良率", "有無添加回收料", "廢蓋KG", "廢料KG", "可回收廢蓋KG", "不良原因",
]

def _extract_rows(data_items, col_xs):
    groups = _group_by_y(data_items, tol=22)  # 手寫同一列跨欄 Y 誤差較大
    sorted_cols = sorted(col_xs.items(), key=lambda x: x[1]) if col_xs else []
    rows = []
    for group in groups:
        row = _map_to_cols(group, sorted_cols)
        if any(v for v in row.values()):
            rows.append(row)
    return rows


def _map_to_cols(group, sorted_cols):
    row = {col: "" for col in ALL_COLS}
    if not sorted_cols:
        for i, item in enumerate(sorted(group, key=lambda i: i["x"])):
            if i < len(ALL_COLS):
                row[ALL_COLS[i]] = item["text"]
        return row

    col_names = [c[0] for c in sorted_cols]
    col_pos   = [c[1] for c in sorted_cols]

    for item in sorted(group, key=lambda i: i["x"]):
        dists = [abs(item["x"] - cx) for cx in col_pos]
        idx = dists.index(min(dists))
        col_key = col_names[idx]
        text = item["text"]
        if col_key == "機台編號":
            text = _fix_machine_number(text)
        if col_key == "不良率" and text and "%" not in text and re.search(r"\d", text):
            text += "%"
        # 孤立的 "%" 是被 Vision 拆開的不良率符號，補回不良率欄
        if text == "%" and col_key != "不良率":
            if "%" not in row["不良率"]:
                row["不良率"] = (row["不良率"] + "%").strip()
            continue
        row[col_key] = (row[col_key] + text).strip()
    return row


def _fix_machine_number(text):
    text = re.sub(r"\s", "", text).upper()
    n = text.replace("O", "0").replace("I", "1").replace("L", "1").replace("G", "6")
    m = re.match(r"B[C0](\d+)", n)
    if m:
        return f"BC{m.group(1).zfill(2)}"
    digits = re.findall(r"\d+", n)
    return f"BC{''.join(digits).zfill(2)}" if digits else text


# ── footer ─────────────────────────────────────────────────────────────────────

def _extract_footer(items):
    footer = {"廠長": "", "副廠": "", "填表人": ""}
    rows = _group_by_y(items, tol=14)
    for group in rows[-5:]:  # 只看最後幾行
        line = "".join(i["text"] for i in group)
        for key in footer:
            if key in line:
                m = re.search(rf"{key}:?\s*(\S+)", line)
                if m:
                    val = m.group(1)
                    if val not in ["副廠", "廠長", "填表人"]:
                        footer[key] = val
    return footer


# ── helpers ────────────────────────────────────────────────────────────────────

def _group_by_y(items, tol=14):
    if not items:
        return []
    groups, cur = [], [items[0]]
    ref_y = items[0]["y"]
    for item in items[1:]:
        if abs(item["y"] - ref_y) <= tol:
            cur.append(item)
        else:
            groups.append(sorted(cur, key=lambda i: i["x"]))
            cur, ref_y = [item], item["y"]
    groups.append(sorted(cur, key=lambda i: i["x"]))
    return groups
