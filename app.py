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

# ---- Timezone VN ----
VN_TZ = timezone(timedelta(hours=7))

# ---- Load MONGO_URI từ biến môi trường ----
MONGO_URI = os.getenv(
    "MONGO_URI",
    "mongodb+srv://banhbaobeo2205:lm2hiCLXp6B0D7hq@cluster0.festnla.mongodb.net/?retryWrites=true&w=majority"
)
DB_NAME = os.getenv("DB_NAME", "Sun_Database_1")

if not MONGO_URI or MONGO_URI.strip() == "":
    raise ValueError("❌ Lỗi: MONGO_URI chưa được cấu hình trong biến môi trường Render!")

# ---- Kết nối MongoDB ----
try:
    client = MongoClient(MONGO_URI)
    db = client[DB_NAME]
    collection = db["alt_checkins"]
except Exception as e:
    raise RuntimeError(f"❌ Không thể kết nối MongoDB: {e}")


# ---- Helper function ----
def build_query(filter_type, start_date, end_date, search):
    query = {}
    today = datetime.now()

    # Nếu có khoảng ngày custom
    if start_date and end_date:
        try:
            start = datetime.strptime(start_date, "%Y-%m-%d")
            end = datetime.strptime(end_date, "%Y-%m-%d") + timedelta(days=1) - timedelta(seconds=1)
            query["CheckinTime"] = {"$gte": start, "$lte": end}
        except ValueError:
            pass

    # Lọc nhanh theo tuần/tháng/năm
    elif filter_type == "week":
        start = today - timedelta(days=today.weekday())  # Monday
        end = start + timedelta(days=6)  # Sunday
        query["CheckinTime"] = {"$gte": start, "$lte": end}

    elif filter_type == "month":
        start = today.replace(day=1)
        last_day = calendar.monthrange(today.year, today.month)[1]
        end = today.replace(day=last_day, hour=23, minute=59, second=59)
        query["CheckinTime"] = {"$gte": start, "$lte": end}

    elif filter_type == "year":
        start = today.replace(month=1, day=1)
        end = today.replace(month=12, day=31, hour=23, minute=59, second=59)
        query["CheckinTime"] = {"$gte": start, "$lte": end}

    # Nếu có tham số tìm kiếm
    if search:
        query["EmployeeName"] = {"$regex": re.compile(search, re.IGNORECASE)}

    return query


# ---- API routes ----
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/attendances", methods=["GET"])
def get_attendances():
    try:
        filter_type = request.args.get("filter", "all")
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        query = build_query(filter_type, start_date, end_date, search)

        data = list(collection.find(query, {
            "_id": 0,
            "EmployeeId": 1,
            "EmployeeName": 1,
            "Address": 1,
            "CheckinTime": 1,
            "Shift": 1,
            "Status": 1,
            "FaceImage": 1   # thêm cột ảnh
        }))

        # Convert datetime -> string theo giờ VN (dd/MM/yyyy HH:mm:ss)
        for d in data:
            if isinstance(d.get("CheckinTime"), datetime):
                d["CheckinTime"] = d["CheckinTime"].astimezone(VN_TZ).strftime("%d/%m/%Y %H:%M:%S")

        return jsonify(data), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/export-excel", methods=["GET"])
def export_to_excel():
    try:
        filter_type = request.args.get("filter", "all")
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")
        search = request.args.get("search", "").strip()

        query = build_query(filter_type, start_date, end_date, search)
        data = list(collection.find(query, {
            "_id": 0,
            "EmployeeId": 1,
            "EmployeeName": 1,
            "Address": 1,
            "CheckinTime": 1,
            "Shift": 1,
            "Status": 1
        }))

        # Convert datetime -> string theo giờ VN (dd/MM/yyyy HH:mm:ss)
        for d in data:
            if isinstance(d.get("CheckinTime"), datetime):
                d["CheckinTime"] = d["CheckinTime"].astimezone(VN_TZ).strftime("%d/%m/%Y %H:%M:%S")

        df = pd.DataFrame(data)
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Attendances", index=False)
        output.seek(0)

        # Tạo tên file động
        if start_date and end_date:
            filename = f"Danh sách chấm công_{start_date}_to_{end_date}.xlsx"
        elif search:
            filename = f"Danh sách chấm công_{search}_{datetime.now().strftime('%d/%m/%Y')}.xlsx"
        else:
            filename = f"Danh sách chấm công_{filter_type}_{datetime.now().strftime('%d/%m/%Y')}.xlsx"

        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
