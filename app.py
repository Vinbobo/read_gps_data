from flask import Flask, render_template, jsonify, send_file, request
from pymongo import MongoClient
from flask_cors import CORS
import os
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta, timezone
import calendar
import re

app = Flask(__name__, template_folder="templates")
CORS(app)

VN_TZ = timezone(timedelta(hours=7))

MONGO_URI = os.getenv(
    "MONGO_URI",
    "mongodb+srv://banhbaobeo2205:lm2hiCLXp6B0D7hq@cluster0.festnla.mongodb.net/?retryWrites=true&w=majority"
)
DB_NAME = os.getenv("DB_NAME", "Sun_Database_1")

client = MongoClient(MONGO_URI)
db = client[DB_NAME]
collection = db["alt_checkins"]
idx_collection = db["idx_collection"]   # ✅ dùng để kiểm tra quyền

# Danh sách EmployeeId được phép
ALLOWED_IDS = {"A000", "A001", "A002","A003"}


@app.route("/login", methods=["GET"])
def login():
    emp_id = request.args.get("empId")
    if not emp_id:
        return jsonify({"success": False, "message": "❌ Bạn cần nhập EmployeeId"}), 400

    # Kiểm tra trong danh sách cho phép
    if emp_id in ALLOWED_IDS:
        return jsonify({"success": True, "message": "✅ Đăng nhập thành công"})
    else:
        return jsonify({"success": False, "message": "🚫 EmployeeId không có quyền truy cập"}), 403


@app.route("/")
def index():
    return render_template("index.html")


# ---- Query builder ----
def build_query(filter_type, start_date, end_date, search, shift=None):
    query = {}
    today = datetime.now(VN_TZ)

    if filter_type == "custom" and start_date and end_date:
        query["CheckinDate"] = {"$gte": start_date, "$lte": end_date}
    elif filter_type == "week":
        start = (today - timedelta(days=today.weekday())).strftime("%Y-%m-%d")
        end = (today + timedelta(days=6 - today.weekday())).strftime("%Y-%m-%d")
        query["CheckinDate"] = {"$gte": start, "$lte": end}
    elif filter_type == "month":
        start = today.replace(day=1).strftime("%Y-%m-%d")
        last_day = calendar.monthrange(today.year, today.month)[1]
        end = today.replace(day=last_day).strftime("%Y-%m-%d")
        query["CheckinDate"] = {"$gte": start, "$lte": end}
    elif filter_type == "year":
        start = today.replace(month=1, day=1).strftime("%Y-%m-%d")
        end = today.replace(month=12, day=31).strftime("%Y-%m-%d")
        query["CheckinDate"] = {"$gte": start, "$lte": end}

    if search:
        query["EmployeeId"] = {"$regex": re.compile(search, re.IGNORECASE)}

    if shift:
        if shift.lower() == "sang":
            query["Shift"] = {"$regex": re.compile("Ca 1", re.IGNORECASE)}
        elif shift.lower() == "chieu":
            query["Shift"] = {"$regex": re.compile("Ca 2", re.IGNORECASE)}
        else:
            query["Shift"] = {"$regex": re.compile(shift, re.IGNORECASE)}

    return query


@app.route("/api/attendances", methods=["GET"])
def get_attendances():
    emp_id = request.args.get("empId")
    if not emp_id or emp_id not in ALLOWED_IDS:
        return jsonify({"error": "🚫 Bạn cần cung cấp EmployeeId để truy cập."}), 403

    filter_type = request.args.get("filter", "all")
    start_date = request.args.get("startDate")
    end_date = request.args.get("endDate")
    search = request.args.get("search", "").strip()
    shift = request.args.get("shift")

    query = build_query(filter_type, start_date, end_date, search, shift)

    data = list(collection.find(query, {
        "_id": 0,
        "EmployeeId": 1,
        "EmployeeName": 1,
        "ProjectId": 1,
        "Tasks": 1,
        "OtherNote": 1,
        "Address": 1,
        "CheckinTime": 1,
        "Shift": 1,
        "Status": 1,
        "FaceImage": 1
    }))

    for d in data:
        if isinstance(d.get("CheckinTime"), datetime):
            d["CheckinTime"] = d["CheckinTime"].astimezone(VN_TZ).strftime("%d/%m/%Y %H:%M:%S")

    return jsonify(data)


@app.route("/api/export-excel", methods=["GET"])
def export_to_excel():
    emp_id = request.args.get("empId")
    if not emp_id or emp_id not in ALLOWED_IDS:
        return jsonify({"error": "🚫 Bạn cần cung cấp EmployeeId để xuất Excel."}), 403

    filter_type = request.args.get("filter", "all")
    start_date = request.args.get("startDate")
    end_date = request.args.get("endDate")
    search = request.args.get("search", "").strip()
    shift = request.args.get("shift")

    query = build_query(filter_type, start_date, end_date, search, shift)
    data = list(collection.find(query, {
        "_id": 0,
        "EmployeeId": 1,
        "EmployeeName": 1,
        "ProjectId": 1,
        "Tasks": 1,
        "OtherNote": 1,
        "Address": 1,
        "CheckinTime": 1,
        "Shift": 1,
        "Status": 1
    }))

    for d in data:
        if isinstance(d.get("CheckinTime"), datetime):
            d["CheckinTime"] = d["CheckinTime"].astimezone(VN_TZ).strftime("%d/%m/%Y %H:%M:%S")
        if isinstance(d.get("Tasks"), list):
            d["Tasks"] = ", ".join(d["Tasks"])

    df = pd.DataFrame(data)
    df.rename(columns={
        "EmployeeId": "Mã NV",
        "EmployeeName": "Tên nhân viên",
        "ProjectId": "Mã dự án",
        "Tasks": "Công việc",
        "OtherNote": "Khác",
        "Address": "Địa chỉ",
        "CheckinTime": "Thời gian Check-in",
        "Shift": "Ca làm việc",
        "Status": "Trạng thái"
    }, inplace=True)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Chấm công", index=False)
    output.seek(0)

    filename = f"ChamCong_{filter_type}_{datetime.now().strftime('%d-%m-%Y')}.xlsx"
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
