import os
import uuid
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from ocr_parser import parse_form
from form_template import create_blank_template, fill_template
from image_preprocess import enhance

app = Flask(__name__)

ALLOWED_ORIGINS = [
    "https://markcqhsu.github.io",
    "http://localhost:8080",
    "http://127.0.0.1:8080",
]
CORS(app, resources={r"/*": {"origins": ALLOWED_ORIGINS}})

UPLOAD_DIR = "/tmp/uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

ALLOWED_EXTENSIONS = {"jpg", "jpeg", "png", "bmp", "webp"}

def _allowed(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


@app.route("/debug", methods=["POST"])
def debug():
    """回傳 Vision API 原始擷取的 items，用於診斷欄位對應問題"""
    if "image" not in request.files:
        return jsonify({"error": "No image"}), 400
    file = request.files["image"]
    if not file.filename or not _allowed(file.filename):
        return jsonify({"error": "不支援的圖片格式"}), 400
    ext  = file.filename.rsplit(".", 1)[-1].lower()
    tmp  = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}.{ext}")
    file.save(tmp)
    try:
        from ocr_parser_google import _extract_items, get_client
        from google.cloud import vision as v
        with open(tmp, "rb") as f:
            content = f.read()
        image    = v.Image(content=content)
        response = get_client().document_text_detection(image=image)
        items    = _extract_items(response)
        items.sort(key=lambda i: (round(i["y"] / 10), i["x"]))
        from ocr_parser_google import _detect_columns, _extract_meta
        meta = _extract_meta(items)
        col_xs, h_y = _detect_columns(items)
        return jsonify({
            "meta": meta,
            "header_y": h_y,
            "col_xs": col_xs,
            "meta_area": [i for i in items if i["y"] < 160],
        })
    finally:
        os.remove(tmp)


@app.route("/ocr", methods=["POST"])
def ocr():
    if "image" not in request.files:
        return jsonify({"error": "請上傳圖片 (field: image)"}), 400

    file = request.files["image"]
    if not file.filename or not _allowed(file.filename):
        return jsonify({"error": "不支援的圖片格式"}), 400

    form_type = request.form.get("form_type", "cap")
    ext       = file.filename.rsplit(".", 1)[1].lower()
    tmp_path  = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}.{ext}")
    file.save(tmp_path)
    enhance(tmp_path)

    try:
        if form_type == "husky":
            from ocr_parser_husky import parse_form as parse_husky
            result = parse_husky(tmp_path)
        elif form_type == "hpp5":
            from ocr_parser_hpp5 import parse_form as parse_hpp5
            result = parse_hpp5(tmp_path)
        elif form_type == "netstal":
            from ocr_parser_netstal import parse_form as parse_netstal
            result = parse_netstal(tmp_path)
        else:
            result = parse_form(tmp_path)
    finally:
        os.remove(tmp_path)

    return jsonify(result)


@app.route("/template", methods=["GET"])
def template():
    form_type = request.args.get("type", "cap")
    if form_type == "husky":
        from form_template_husky import create_blank_template as blank_husky
        machine = request.args.get("machine", "EM04")
        path = blank_husky(machine)
        return send_file(
            path, as_attachment=True,
            download_name=f"Husky_製程管制標準_{machine}_空白.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    if form_type == "hpp5":
        from form_template_hpp5 import create_blank_template as blank_hpp5
        path = blank_hpp5()
        return send_file(
            path, as_attachment=True,
            download_name="Hpp5_製程管制標準_空白.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    if form_type == "netstal":
        from form_template_netstal import create_blank_template as blank_netstal
        path = blank_netstal()
        return send_file(
            path, as_attachment=True,
            download_name="Netstal_製程管制標準_空白.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    process = request.args.get("process", "切割")
    path = create_blank_template(process)
    return send_file(
        path,
        as_attachment=True,
        download_name="塑蓋廠不良率紀錄表_blank.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/export", methods=["POST"])
def export():
    data = request.get_json()
    if not data:
        return jsonify({"error": "請提供 JSON 資料"}), 400

    form_type = data.get("form_type", "cap")
    if form_type == "husky":
        from form_template_husky import fill_template as fill_husky
        path = fill_husky(data)
        return send_file(
            path,
            as_attachment=True,
            download_name="Husky_製程管制標準.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    if form_type == "hpp5":
        from form_template_hpp5 import fill_template as fill_hpp5
        path = fill_hpp5(data)
        return send_file(
            path,
            as_attachment=True,
            download_name="Hpp5_製程管制標準.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    if form_type == "netstal":
        from form_template_netstal import fill_template as fill_netstal
        path = fill_netstal(data)
        return send_file(
            path,
            as_attachment=True,
            download_name="Netstal_製程管制標準.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    path = fill_template(data)
    process = data.get("製程", "不良率紀錄")
    return send_file(
        path,
        as_attachment=True,
        download_name=f"{process}_不良率紀錄表.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)
