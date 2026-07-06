from __future__ import annotations

import importlib.util
import os
from pathlib import Path

from flask import Flask, jsonify, render_template, request

BASE_DIR = Path(__file__).resolve().parent
SCRIPT_PATH = BASE_DIR / "提取脚本.py"
_extractor = None
_extractor_mtime_ns = None

VISIBLE_COLUMNS = [
    "交易方向",
    "交易日期",
    "债券代码",
    "债券简称",
    "到期收益率",
    "行权收益率",
    "原始净价",
    "交易净价",
    "交易规模万",
    "我方账户",
    "对方账户",
    "过券",
    "中介",
    "中介费",
    "对手方交易员",
    "清算速度",
    "约定号",
    "对手方交易员代码",
    "对手方交易商代码",
    "对手方交易商简称",
    "对手方交易主体代码",
    "报价发起方",
]


def load_extractor():
    spec = importlib.util.spec_from_file_location("bond_extractor", SCRIPT_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法加载提取脚本: {SCRIPT_PATH}")

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def get_extractor():
    global _extractor, _extractor_mtime_ns

    current_mtime_ns = SCRIPT_PATH.stat().st_mtime_ns
    if _extractor is None or _extractor_mtime_ns != current_mtime_ns:
        _extractor = load_extractor()
        _extractor_mtime_ns = current_mtime_ns

    return _extractor


app = Flask(__name__)


@app.after_request
def add_local_access_headers(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    return response


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/healthz")
def healthz():
    return jsonify(
        {
            "ok": True,
            "script": SCRIPT_PATH.name,
            "script_exists": SCRIPT_PATH.exists(),
        }
    )


@app.route("/api/recognize-agent", methods=["POST", "OPTIONS"])
def recognize_agent():
    if request.method == "OPTIONS":
        return ("", 204)

    payload = request.get_json(silent=True) or {}
    text = (payload.get("text") or payload.get("input_text") or "").strip()
    if not text:
        return jsonify({"success": False, "message": "请输入待识别文本。"}), 400

    try:
        extractor = get_extractor()
        rows = extractor.parse_text(text)
    except Exception as exc:
        return (
            jsonify(
                {
                    "success": False,
                    "message": f"识别脚本执行失败: {exc}",
                }
            ),
            500,
        )

    normalized_rows = [
        {column: row.get(column, "") for column in extractor.OUTPUT_COLUMNS}
        for row in rows
    ]
    return jsonify(
        {
            "success": True,
            "count": len(normalized_rows),
            "message": "识别完成" if normalized_rows else "未识别到可解析记录",
            "columns": VISIBLE_COLUMNS,
            "rows": normalized_rows,
        }
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    host = os.environ.get("HOST", "0.0.0.0")
    app.run(host=host, port=port, debug=False, use_reloader=False)
