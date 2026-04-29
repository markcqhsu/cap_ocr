import base64
import json
import os
import re

import anthropic

_client = None

def get_client():
    global _client
    if _client is None:
        _client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
    return _client


PROMPT = """這是一張塑蓋廠各站別不良率紀錄表的手寫照片。

請仔細辨識圖片中的所有文字，並以 JSON 格式回傳以下資料（只回傳 JSON，不要其他說明）：

{
  "製程": "例如：切割、射出、組裝、包裝、印刷",
  "組別": "數字",
  "班別": "英文字母，例如 A、B、C",
  "年份": "四位數年份",
  "rows": [
    {
      "日期": "月/日 格式，同一機台多筆時第一筆有日期，後續可為空",
      "機台編號": "BC 開頭加數字，例如 BC09",
      "工單號碼": "五位數字",
      "品名蓋型": "產品名稱",
      "產量": "數字",
      "不良數": "數字",
      "不良率": "百分比，包含 % 符號",
      "有無添加回收料": "有 或 無，若空白填空字串",
      "廢蓋KG": "數字，若空白填空字串",
      "廢料KG": "數字，若空白填空字串",
      "可回收廢蓋KG": "數字，若空白填空字串",
      "不良原因": "文字，若空白填空字串"
    }
  ],
  "廠長": "簽名文字，若空白填空字串",
  "副廠": "簽名文字，若空白填空字串",
  "填表人": "簽名文字，若空白填空字串"
}

注意：
- 機台編號統一格式為 BC + 兩位數，例如 BC09、BC11、BC16
- 不良率數字後記得加 %
- 表格中同一機台可能有多筆工單，每筆工單獨立一行
- 手寫字可能不清晰，請盡力辨識最接近的正確值"""


def parse_form(image_path: str) -> dict:
    with open(image_path, "rb") as f:
        image_data = base64.standard_b64encode(f.read()).decode("utf-8")

    ext = image_path.rsplit(".", 1)[-1].lower()
    media_type_map = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png", "bmp": "image/bmp", "webp": "image/webp"}
    media_type = media_type_map.get(ext, "image/jpeg")

    message = get_client().messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=2048,
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

    # 去除可能的 markdown code block
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        # 嘗試擷取 JSON 區塊
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        if m:
            return json.loads(m.group())
        return _empty()


def _empty():
    return {
        "製程": "", "組別": "", "班別": "", "年份": "",
        "rows": [],
        "廠長": "", "副廠": "", "填表人": "",
    }
