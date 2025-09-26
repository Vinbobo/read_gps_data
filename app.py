from flask import Flask, render_template, jsonify, send_file, request, session, redirect, url_for
from pymongo import MongoClient
from flask_cors import CORS
import os
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta, timezone
import calendar
import re

app = Flask(__name__, template_folder="templates")
app.secret_key = os.getenv("SECRET_KEY", "supersecret")  # cần để dùng session
CORS(app)

# ---- Timezone VN ----
VN_TZ = timezone(timedelta(hours=7))

# ---- Load MONGO_URI ----
MONGO_URI = os.getenv(
    "MONGO_URI",
    "mongodb+srv://banhbaobeo2205:lm2hiCLXp6B0D7hq@cluster0.festnla.mongodb.net/?retryWrites=true&w=majority"
)
DB_NAME = os.getenv("DB_NAME", "Sun_Database_1")

if not MONGO_URI or MONGO_URI.strip() == "":
    raise ValueError("❌ Lỗi: MONGO_URI chưa được cấu hình!")

# ---- Kết nối MongoDB ----
try:
    client = MongoClient(MONGO_URI)
    db = client[DB_NAME]
    collection = db["alt_checkins"]
    idx_collection = db["idx_collection"]   # ✅ nơi chứa danh sách nhân viên
except Exception as e:
    raise RuntimeError(f"❌ Không thể kết nối MongoDB: {e}")


# ---- Danh sách EmployeeId có quyền xem ----
AUTHORIZED_IDS = {"S002", "S018", "S019"}


# ---- Middleware kiểm tra quyền ----
@app.before_request
def restrict_access():
    if request.endpoint in ["index", "get_attendances", "export_to_excel"]:
        # Lấy EmployeeId từ session hoặc query string
        emp_id = session.get("EmployeeId") or request.args.get("empId")

        if not emp_id:
            return "🚫 Bạn cần cung cấp EmployeeId để truy cập.", 403

        # Kiểm tra trong idx_collection có tồn tại không
        user = idx_collection.find_one({"EmployeeId": emp_id})
        if not user:
            return "🚫 EmployeeId không tồn tại trong hệ thống.", 403

        # Kiểm tra quyền
        if emp_id not in AUTHORIZED_IDS:
            return "🚫 Bạn không có quyền truy cập trang này.", 403

        # ✅ Nếu pass → lưu vào session
        session["EmployeeId"] = emp_id


# ---- Xây dựng query cho filter ----
def build_query(filter_type, start_date, end_date, search, shift=None):
    query = {}
    today = datetime.now(VN_TZ)

    # ---- Lọc theo thời gian ----
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

    # ---- Lọc theo EmployeeId ----
    if search:
        query["EmployeeId"] = {"$regex": re.compile(search, re.IGNORECASE)}

    # ---- Lọc theo ca ----
    if shift:
        if shift.lower() == "sang":
            query["Shift"] = {"$regex": re.compile("Ca 1", re.IGNORECASE)}
        elif shift.lower() == "chieu":
            query["Shift"] = {"$regex": re.compile("Ca 2", re.IGNORECASE)}
        else:
            query["Shift"] = {"$regex": re.compile(shift, re.IGNORECASE)}

    return query


# ---- API: Trang index ----
@app.route("/")
def index():
    return render_template("index.html")


# ---- API: Lấy danh sách chấm công ----
@app.route("/api/attendances", methods=["GET"])
def get_attendances():
    try:
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

        # Convert datetime -> string
        for d in data:
            if isinstance(d.get("CheckinTime"), datetime):
                d["CheckinTime"] = d["CheckinTime"].astimezone(VN_TZ).strftime("%d/%m/%Y %H:%M:%S")

        return jsonify(data), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ---- API: Xuất Excel ----
@app.route("/api/export-excel", methods=["GET"])
def export_to_excel():
    try:
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

        return send_file(
            output,
            as_attachment=True,
            download_name=f"Danh_sach_cham_cong.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
