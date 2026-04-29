import tempfile
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from ocr_parser_husky import LEFT_TEMPLATE, RIGHT_TEMPLATE


def _thin():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

def _center(wrap=False):
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)

def _left():
    return Alignment(horizontal="left", vertical="center")

HEADER_FILL = PatternFill("solid", fgColor="1A3C6E")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=9)
GROUP_FILL  = PatternFill("solid", fgColor="D9E4F0")
GROUP_FONT  = Font(bold=True, size=9)
DATA_FONT   = Font(size=9)
TITLE_FONT  = Font(bold=True, size=12)


def create_blank_template(machine_type="Husky") -> str:
    """Generate a blank Husky 製程管制標準 sheet for download/review."""
    params = [{"group": g, "name": n, "value": "", "range": r, "unit": u}
              for g, n, r, u in LEFT_TEMPLATE + RIGHT_TEMPLATE]
    return _build_excel({
        "機型": machine_type, "日期": "", "克重": "", "瓶口": "", "模號": "", "原料": "",
        "params": params,
    }, title=f"Husky 製程管制標準（{machine_type}）空白表格")


def fill_template(data: dict) -> str:
    return _build_excel(data, title="Husky 製程管制標準")


def _build_excel(data: dict, title: str) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "製程管制標準"

    # Column widths: A-E (left), F (gap), G-K (right)
    for col, w in zip("ABCDEFGHIJK", [9, 16, 8, 7, 8, 1, 9, 16, 8, 7, 8]):
        ws.column_dimensions[col].width = w

    # Title
    ws.merge_cells("A1:K1")
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = _center()
    ws.row_dimensions[1].height = 18

    # Meta
    ws.merge_cells("A2:K2")
    ws["A2"] = "    ".join([
        f'機型：{data.get("機型","")}', f'日期：{data.get("日期","")}',
        f'克重：{data.get("克重","")}', f'瓶口：{data.get("瓶口","")}',
        f'模號：{data.get("模號","")}', f'原料：{data.get("原料","")}',
    ])
    ws["A2"].font = Font(size=9)
    ws["A2"].alignment = _left()
    ws.row_dimensions[2].height = 14

    # Section headers (row 3)
    for col, label in zip(list("ABCDE") + list("GHIJK"),
                          ["項目","名稱","設定值","範圍","單位"] * 2):
        c = ws[f"{col}3"]
        c.value = label; c.font = HEADER_FONT; c.fill = HEADER_FILL
        c.alignment = _center(); c.border = _thin()
    ws.row_dimensions[3].height = 14

    # Split params
    right_groups = {"機器加熱設定", "模具加熱設定", "控制", "乾燥機"}
    params = data.get("params", [])
    left_params  = [p for p in params if p["group"] not in right_groups]
    right_params = [p for p in params if p["group"]     in right_groups]

    max_rows = max(len(left_params), len(right_params))

    for idx in range(max_rows):
        er = 4 + idx
        ws.row_dimensions[er].height = 13

        if idx < len(left_params):
            _write_row(ws, er, "A","B","C","D","E", left_params[idx])
        else:
            for c in "ABCDE": ws[f"{c}{er}"].border = _thin()

        if idx < len(right_params):
            _write_row(ws, er, "G","H","I","J","K", right_params[idx])
        else:
            for c in "GHIJK": ws[f"{c}{er}"].border = _thin()

    _merge_groups(ws, left_params,  4, "A")
    _merge_groups(ws, right_params, 4, "G")

    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    wb.save(tmp.name)
    return tmp.name


def _write_row(ws, row, cg, cn, cv, cr, cu, param):
    for c in [cg, cn, cv, cr, cu]:
        ws[f"{c}{row}"].border = _thin()
        ws[f"{c}{row}"].font   = DATA_FONT
    ws[f"{cg}{row}"].value     = param["group"]
    ws[f"{cg}{row}"].fill      = GROUP_FILL
    ws[f"{cg}{row}"].font      = GROUP_FONT
    ws[f"{cg}{row}"].alignment = _center(wrap=True)
    ws[f"{cn}{row}"].value     = param["name"];  ws[f"{cn}{row}"].alignment = _left()
    ws[f"{cv}{row}"].value     = param["value"]; ws[f"{cv}{row}"].alignment = _center()
    ws[f"{cr}{row}"].value     = param["range"]; ws[f"{cr}{row}"].alignment = _center()
    ws[f"{cu}{row}"].value     = param["unit"];  ws[f"{cu}{row}"].alignment = _center()


def _merge_groups(ws, params, start_row, col):
    if not params:
        return
    i = 0
    while i < len(params):
        group = params[i]["group"]
        j = i + 1
        while j < len(params) and params[j]["group"] == group:
            j += 1
        r1, r2 = start_row + i, start_row + j - 1
        if r2 > r1:
            ws.merge_cells(f"{col}{r1}:{col}{r2}")
        c = ws[f"{col}{r1}"]
        c.value = group; c.fill = GROUP_FILL; c.font = GROUP_FONT
        c.alignment = _center(wrap=True); c.border = _thin()
        i = j
