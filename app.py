from __future__ import annotations

import os
import threading
import webbrowser
from io import BytesIO
from pathlib import Path
from urllib.parse import quote

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.exceptions import RequestEntityTooLarge

from normalizer import build_workbook, extract_records


app = Flask(__name__)
# Vercel Functions의 요청/응답 본문 제한(4.5MB)보다 여유 있게 제한한다.
app.config["MAX_CONTENT_LENGTH"] = 4 * 1024 * 1024


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/convert")
def convert():
    upload = request.files.get("file")
    if upload is None or not upload.filename:
        return jsonify(error="변환할 엑셀 파일을 선택해 주세요."), 400

    source_name = Path(upload.filename).name
    if Path(source_name).suffix.lower() != ".xlsx":
        return jsonify(error=".xlsx 형식의 엑셀 파일만 사용할 수 있습니다."), 400

    try:
        source_bytes = BytesIO(upload.read())
        records, physical_rows, merged_fragments = extract_records(source_bytes)
        if not records:
            return jsonify(
                error="학생 기록을 찾지 못했습니다. 학교생활기록부 조회 파일인지 확인해 주세요."
            ), 422
        output_bytes = build_workbook(records)
    except (ValueError, KeyError) as exc:
        return jsonify(error=str(exc)), 422
    except Exception:
        app.logger.exception("Excel conversion failed")
        return jsonify(
            error="파일을 처리하지 못했습니다. 파일이 손상되지 않았는지 확인해 주세요."
        ), 500

    download_name = f"{Path(source_name).stem}_정리.xlsx"
    response = send_file(
        output_bytes,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        max_age=0,
    )
    response.headers["X-Record-Count"] = str(len(records))
    response.headers["X-Physical-Row-Count"] = str(physical_rows)
    response.headers["X-Merged-Count"] = str(merged_fragments)
    response.headers["X-Download-Name"] = quote(download_name)
    response.headers["Cache-Control"] = "no-store"
    return response


@app.errorhandler(RequestEntityTooLarge)
def too_large(_error):
    return jsonify(error="파일 크기는 4MB 이하여야 합니다."), 413


def open_browser() -> None:
    webbrowser.open("http://127.0.0.1:8765")


if __name__ == "__main__":
    if os.environ.get("OPEN_BROWSER") == "1":
        threading.Timer(1.2, open_browser).start()
    app.run(host="127.0.0.1", port=8765, debug=False)
