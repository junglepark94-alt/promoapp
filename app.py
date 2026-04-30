from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for
import json
import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io

app = Flask(__name__)

DATA_FILE = "data.json"

TEAMS = [
    "냉동BM1팀", "냉동BM2팀", "발효유팀", "Dairy제품팀",
    "커피 음료팀", "NC팀", "해외", "콘텐츠전략팀"
]

CATEGORIES = ["신제품", "리뉴얼", "프로모션", "박람회"]

SCHEDULE_OPTIONS = [
    "1월 초", "1월 중", "1월 말", "1월(미정)",
    "2월 초", "2월 중", "2월 말", "2월(미정)",
    "3월 초", "3월 중", "3월 말", "3월(미정)",
    "4월 초", "4월 중", "4월 말", "4월(미정)",
    "5월 초", "5월 중", "5월 말", "5월(미정)",
    "6월 초", "6월 중", "6월 말", "6월(미정)",
    "7월 초", "7월 중", "7월 말", "7월(미정)",
    "8월 초", "8월 중", "8월 말", "8월(미정)",
    "9월 초", "9월 중", "9월 말", "9월(미정)",
    "10월 초", "10월 중", "10월 말", "10월(미정)",
    "11월 초", "11월 중", "11월 말", "11월(미정)",
    "12월 초", "12월 중", "12월 말", "12월(미정)",
]

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

@app.route("/")
def index():
    return render_template("index.html", teams=TEAMS, categories=CATEGORIES, schedule_options=SCHEDULE_OPTIONS)

@app.route("/admin")
def admin():
    return redirect(url_for("index"))

@app.route("/api/submit", methods=["POST"])
def submit():
    body = request.json
    entries = body.get("entries", [])
    if not entries:
        return jsonify({"success": False, "message": "데이터가 없습니다."}), 400

    data = load_data()
    team = entries[0].get("team", "")
    # Remove existing entries from this team submission (by submitted_at grouping)
    # and append new ones
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for entry in entries:
        entry["submitted_at"] = now
        entry["id"] = f"{now}_{len(data)}"
        data.append(entry)

    save_data(data)
    return jsonify({"success": True, "message": f"{len(entries)}건 제출 완료!"})

@app.route("/api/data")
def get_data():
    data = load_data()
    team = request.args.get("team")
    if team:
        data = [d for d in data if d.get("team") == team]
    return jsonify(data)

@app.route("/api/delete/<entry_id>", methods=["DELETE"])
def delete_entry(entry_id):
    data = load_data()
    data = [d for d in data if d.get("id") != entry_id]
    save_data(data)
    return jsonify({"success": True})

@app.route("/api/clear_team", methods=["POST"])
def clear_team():
    team = request.json.get("team")
    data = load_data()
    data = [d for d in data if d.get("team") != team]
    save_data(data)
    return jsonify({"success": True})

@app.route("/export")
def export():
    data = load_data()

    wb = openpyxl.Workbook()

    # ── 스타일 정의 ──────────────────────────────────────────
    header_font = Font(name="맑은 고딕", bold=True, size=10, color="FFFFFF")
    body_font   = Font(name="맑은 고딕", size=10)
    title_font  = Font(name="맑은 고딕", bold=True, size=13)

    navy   = PatternFill("solid", fgColor="1F3864")
    blue   = PatternFill("solid", fgColor="2E75B6")
    lt_blue= PatternFill("solid", fgColor="DEEAF1")
    white  = PatternFill("solid", fgColor="FFFFFF")
    stripe = PatternFill("solid", fgColor="F2F7FB")

    thin = Side(style="thin", color="B0C4D8")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    COLS  = ["일정", "담당팀", "담당자", "구분", "제품명", "비고"]
    WIDTHS = [14, 14, 10, 10, 32, 40]

    def make_sheet(ws, rows, sheet_title):
        ws.row_dimensions[1].height = 22
        ws.row_dimensions[2].height = 16
        ws.row_dimensions[3].height = 16
        ws.row_dimensions[4].height = 16
        ws.row_dimensions[5].height = 22

        # 타이틀 영역 (A1:F4)
        ws.merge_cells("A1:F1")
        ws["A1"] = "제품 출시 및 프로모션 일정 취합본"
        ws["A1"].font = Font(name="맑은 고딕", bold=True, size=14, color="1F3864")
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
        ws["A1"].fill = white

        ws.merge_cells("A2:F2")
        ws["A2"] = f"작성일자: {datetime.now().strftime('%Y.%m.%d')}    작성팀: 홍보담당 콘텐츠전략팀"
        ws["A2"].font = Font(name="맑은 고딕", size=9, color="595959")
        ws["A2"].alignment = left

        ws.merge_cells("A3:F3")
        ws["A3"] = "※ 아래는 당월 및 향후 2개월 일정으로, 익월 이후 일정은 변동 가능성에 유의"
        ws["A3"].font = Font(name="맑은 고딕", size=9, italic=True, color="C00000")
        ws["A3"].alignment = left

        # 빈 구분선 행
        ws.merge_cells("A4:F4")
        ws["A4"].fill = white

        # 헤더 행
        for c_idx, (col_name, width) in enumerate(zip(COLS, WIDTHS), 1):
            cell = ws.cell(row=5, column=c_idx, value=col_name)
            cell.font   = header_font
            cell.fill   = navy
            cell.alignment = center
            cell.border = border
            ws.column_dimensions[cell.column_letter].width = width
        ws.row_dimensions[5].height = 20

        # 데이터 행
        for r_idx, row in enumerate(rows, 6):
            fill = stripe if r_idx % 2 == 0 else white
            ws.row_dimensions[r_idx].height = 18
            for c_idx, key in enumerate(["schedule", "team", "name", "category", "product", "note"], 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=row.get(key, ""))
                cell.font   = body_font
                cell.fill   = fill
                cell.alignment = center if c_idx <= 4 else left
                cell.border = border

        ws.freeze_panes = "A6"

    # 취합 시트
    ws_all = wb.active
    ws_all.title = "취합"

    MONTH_ORDER = {m: i for i, m in enumerate([
        "1월","2월","3월","4월","5월","6월",
        "7월","8월","9월","10월","11월","12월"
    ])}
    PERIOD_ORDER = {"초": 0, "중": 1, "말": 2, "(미정)": 3}

    def sort_key(d):
        s = d.get("schedule", "")
        for m, mi in MONTH_ORDER.items():
            if s.startswith(m):
                rest = s[len(m):]
                pi = next((v for k, v in PERIOD_ORDER.items() if k in rest), 99)
                return (mi, pi, d.get("team", ""), d.get("name", ""))
        return (99, 99, d.get("team", ""), d.get("name", ""))

    sorted_data = sorted(data, key=sort_key)
    make_sheet(ws_all, sorted_data, "취합")

    # 팀별 시트
    for team in TEAMS:
        team_data = [d for d in sorted_data if d.get("team") == team]
        if not team_data:
            continue
        ws = wb.create_sheet(title=team[:15])
        make_sheet(ws, team_data, team)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    filename = f"프로모션일정_취합본_{datetime.now().strftime('%Y%m%d')}.xlsx"
    return send_file(buf, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    app.run(debug=True, port=5000)
