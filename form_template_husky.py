import os, tempfile
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill


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


def fill_template(data: dict) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "製程管制標準"

    # ── Column widths ──────────────────────────────────────────────────────────
    # Left section: A(項目) B(名稱) C(設定值) D(範圍) E(單位)
    # Separator: F
    # Right section: G(項目) H(名稱) I(設定值) J(範圍) K(單位)
    for col, w in zip("ABCDE FGHIJK", [9, 14, 8, 7, 8, 1, 9, 14, 8, 7, 8]):
        ws.column_dimensions[col].width = w

    # ── Title row ─────────────────────────────────────────────────────────────
    ws.merge_cells("A1:K1")
    ws["A1"] = "Husky 製程管制標準"
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = _center()

    # ── Meta row ──────────────────────────────────────────────────────────────
    ws.merge_cells("A2:K2")
    meta_parts = [
        f'機型：{data.get("機型", "")}',
        f'日期：{data.get("日期", "")}',
        f'克重：{data.get("克重", "")}',
        f'瓶口：{data.get("瓶口", "")}',
        f'模號：{data.get("模號", "")}',
        f'原料：{data.get("原料", "")}',
    ]
    ws["A2"] = "    ".join(meta_parts)
    ws["A2"].font = Font(size=9)
    ws["A2"].alignment = _left()

    # ── Section headers ────────────────────────────────────────────────────────
    for col, label in zip(["A","B","C","D","E","G","H","I","J","K"],
                          ["項目","名稱","設定值","範圍","單位","項目","名稱","設定值","範圍","單位"]):
        c = ws[f"{col}3"]
        c.value = label
        c.font  = HEADER_FONT
        c.fill  = HEADER_FILL
        c.alignment = _center()
        c.border = _thin()

    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 14
    ws.row_dimensions[3].height = 14

    # ── Split params into left / right ─────────────────────────────────────────
    from ocr_parser_husky import LEFT_TEMPLATE, RIGHT_TEMPLATE
    left_names  = {t[1] for t in LEFT_TEMPLATE}
    right_names = {t[1] for t in RIGHT_TEMPLATE}

    params = data.get("params", [])
    left_params  = [p for p in params if p["name"] in left_names]
    right_params = [p for p in params if p["name"] in right_names]

    # Pad to equal length for side-by-side layout
    max_rows = max(len(left_params), len(right_params))

    # ── Write data ─────────────────────────────────────────────────────────────
    prev_left_group  = None
    prev_right_group = None
    group_start_left  = 4
    group_start_right = 4

    for idx in range(max_rows):
        excel_row = 4 + idx
        ws.row_dimensions[excel_row].height = 13

        # Left side
        if idx < len(left_params):
            p = left_params[idx]
            _write_param_row(ws, excel_row, "A", "B", "C", "D", "E", p)
        else:
            for col in "ABCDE":
                ws[f"{col}{excel_row}"].border = _thin()

        # Right side
        if idx < len(right_params):
            p = right_params[idx]
            _write_param_row(ws, excel_row, "G", "H", "I", "J", "K", p)
        else:
            for col in "GHIJK":
                ws[f"{col}{excel_row}"].border = _thin()

    # Merge group cells (A column for left, G column for right)
    _merge_groups(ws, left_params, 4, "A")
    _merge_groups(ws, right_params, 4, "G")

    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    wb.save(tmp.name)
    return tmp.name


def _write_param_row(ws, row, col_grp, col_name, col_val, col_rng, col_unit, param):
    for col in [col_grp, col_name, col_val, col_rng, col_unit]:
        ws[f"{col}{row}"].border = _thin()
        ws[f"{col}{row}"].font   = DATA_FONT

    ws[f"{col_grp}{row}"].value     = param["group"]
    ws[f"{col_grp}{row}"].fill      = GROUP_FILL
    ws[f"{col_grp}{row}"].font      = GROUP_FONT
    ws[f"{col_grp}{row}"].alignment = _center(wrap=True)

    ws[f"{col_name}{row}"].value     = param["name"]
    ws[f"{col_name}{row}"].alignment = _left()

    ws[f"{col_val}{row}"].value     = param["value"]
    ws[f"{col_val}{row}"].alignment = _center()

    ws[f"{col_rng}{row}"].value     = param["range"]
    ws[f"{col_rng}{row}"].alignment = _center()

    ws[f"{col_unit}{row}"].value     = param["unit"]
    ws[f"{col_unit}{row}"].alignment = _center()


def _merge_groups(ws, params, start_row, col):
    """Merge cells in the group column for consecutive same-group rows."""
    if not params:
        return
    i = 0
    while i < len(params):
        group = params[i]["group"]
        j = i + 1
        while j < len(params) and params[j]["group"] == group:
            j += 1
        r1 = start_row + i
        r2 = start_row + j - 1
        if r2 > r1:
            ws.merge_cells(f"{col}{r1}:{col}{r2}")
        c = ws[f"{col}{r1}"]
        c.value     = group
        c.fill      = GROUP_FILL
        c.font      = GROUP_FONT
        c.alignment = _center(wrap=True)
        c.border    = _thin()
        i = j
