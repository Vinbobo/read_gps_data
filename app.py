from flask import Flask, render_template, jsonify, send_file, request
from pymongo import MongoClient
from flask_cors import CORS
import os
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta

app = Flask(__name__, template_folder="templates")
CORS(app)

# 🔹 MongoDB Atlas connection string
MONGO_URI = os.getenv(
    "MONGO_URI",
    "mongodb+srv://banhbaobeo2205:lm2hiCLXp6B0D7hq@cluster0.festnla.mongodb.net/?retryWrites=true&w=majority"
)
DB_NAME = os.getenv("DB_NAME", "Sun_Database_1")

client = MongoClient(MONGO_URI)
db = client[DB_NAME]
collection = db["checkins"]

# --- Hàm build filter Mongo ---
def build_date_filter(period: str):
    now = datetime.now()
    if period == "week":
        start = now - timedelta(days=now.weekday())   # Monday of this week
    elif period == "month":
        start = datetime(now.year, now.month, 1)
    elif period == "year":
        start = datetime(now.year, 1, 1)
    else:
        return {}
    return {"CheckinTime": {"$gte": start.isoformat()}}  # Lưu ý: cần định dạng datetime chuẩn trong DB

# --- Trang chính ---
@app.route("/")
def index():
    return render_template("index.html")

# --- API dữ liệu chấm công ---
@app.route("/api/attendances", methods=["GET"])
def get_attendances():
    try:
        filter_param = request.args.get("filter", "all")
        date_filter = build_date_filter(filter_param)

        query = {}
        if date_filter:
            query.update(date_filter)

        data = list(collection.find(query, {
            "_id": 0,
            "EmployeeId": 1,
            "EmployeeName": 1,
            "Address": 1,
            "CheckinTime": 1,
            "Status": 1
        }))

        return jsonify(data), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# --- API Xuất Excel ---
@app.route("/api/export-excel", methods=["GET"])
def export_to_excel():
    try:
        filter_param = request.args.get("filter", "all")
        date_filter = build_date_filter(filter_param)

        query = {}
        if date_filter:
            query.update(date_filter)

        data = list(collection.find(query, {
            "_id": 0,
            "EmployeeId": 1,
            "EmployeeName": 1,
            "Address": 1,
            "CheckinTime": 1,
            "Status": 1
        }))

        df = pd.DataFrame(data)
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Attendances", index=False)
        output.seek(0)

        return send_file(
            output,
            download_name="attendance_data.xlsx",
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
