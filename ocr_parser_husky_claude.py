import base64
import json
import os
import re

import anthropic

from ocr_parser_husky import LEFT_TEMPLATE, RIGHT_TEMPLATE

_client = None


def get_client():
    global _client
    if _client is None:
        _client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
    return _client


# Build the ordered param list prompt from templates
_ALL_PARAMS = LEFT_TEMPLATE + RIGHT_TEMPLATE
_PARAM_LIST = "\n".join(
    f'    {{"group": "{g}", "name": "{n}", "value": ""}}'
    for g, n, _r, _u in _ALL_PARAMS
)

PROMPT = f"""這是一張 Husky 射出成型機製程管制標準表的手寫照片。

請仔細辨識圖片中的所有手寫數值，特別注意以下常見的手寫辨識陷阱：
- 手寫「2」常被誤看成「z」、「Z」或「>」
- 手寫「7」常被誤看成「>」、「n」或「γ」
- 手寫「1」常被誤看成「l」（小寫L）或「I」
- 手寫「6」在小數點後可能看起來像「b」
- 請用每格旁邊的「範圍（±）」欄輔助判斷數值是否合理

表頭請辨識：
- 克重（數字）
- 瓶口（材料代號，如 PCO）
- 模號（英數字）
- 機型（EM04 / EM06 / EM07，根據勾選的核取方塊）
- 日期（若有）

請以 JSON 格式回傳以下資料，只回傳 JSON，不要其他說明：

{{
  "機型": "",
  "日期": "",
  "克重": "",
  "瓶口": "",
  "模號": "",
  "穴數": "",
  "原料": "",
  "params": [
{_PARAM_LIST}
  ]
}}

注意：params 陣列的順序和 group/name 均已固定，只需填入 value 欄位。若某格無法辨識則填空字串。"""


def parse_form(image_path: str) -> dict:
    with open(image_path, "rb") as f:
        image_data = base64.standard_b64encode(f.read()).decode("utf-8")

    ext = image_path.rsplit(".", 1)[-1].lower()
    media_type = {
        "jpg": "image/jpeg", "jpeg": "image/jpeg",
        "png": "image/png", "bmp": "image/bmp", "webp": "image/webp",
    }.get(ext, "image/jpeg")

    message = get_client().messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": media_type,
                            "data": image_data,
                        },
                    },
                    {"type": "text", "text": PROMPT},
                ],
            }
        ],
    )

    raw = message.content[0].text.strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        if m:
            data = json.loads(m.group())
        else:
            return _empty()

    return _normalize(data)


def _normalize(data: dict) -> dict:
    template_index = {(g, n): (r, u) for g, n, r, u in _ALL_PARAMS}

    raw_params = data.get("params", [])
    if isinstance(raw_params, list):
        # Rebuild in canonical template order, filling any gaps
        by_key = {(p.get("group", ""), p.get("name", "")): p.get("value", "")
                  for p in raw_params if isinstance(p, dict)}
        params = [
            {
                "group": g, "name": n,
                "value": by_key.get((g, n), ""),
                "range": r, "unit": u,
            }
            for g, n, r, u in _ALL_PARAMS
        ]
    else:
        params = [
            {"group": g, "name": n, "value": "", "range": r, "unit": u}
            for g, n, r, u in _ALL_PARAMS
        ]

    return {
        "機型": data.get("機型", ""),
        "日期": data.get("日期", ""),
        "克重": data.get("克重", ""),
        "瓶口": data.get("瓶口", ""),
        "模號": data.get("模號", ""),
        "穴數": data.get("穴數", ""),
        "原料": data.get("原料", ""),
        "params": params,
    }


def _empty() -> dict:
    params = [
        {"group": g, "name": n, "value": "", "range": r, "unit": u}
        for g, n, r, u in _ALL_PARAMS
    ]
    return {
        "機型": "", "日期": "", "克重": "", "瓶口": "",
        "模號": "", "穴數": "", "原料": "",
        "params": params,
    }
