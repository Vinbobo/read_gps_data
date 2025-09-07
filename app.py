from flask import Flask, render_template, jsonify, send_file, request
from pymongo import MongoClient
from flask_cors import CORS
import os
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta

app = Flask(__name__, template_folder="templates")
CORS(app)

MONGO_URI = os.getenv(
    "MONGO_URI",
    "mongodb+srv://banhbaobeo2205:lm2hiCLXp6B0D7hq@cluster0.festnla.mongodb.net/?retryWrites=true&w=majority"
)
DB_NAME = os.getenv("DB_NAME", "Sun_Database_1")

client = MongoClient(MONGO_URI)
db = client[DB_NAME]
collection = db["checkins"]

def filter_dataframe(df: pd.DataFrame, period: str) -> pd.DataFrame:
    """Lọc theo tuần / tháng / năm dựa vào cột CheckinTime"""
    if "CheckinTime" not in df.columns:
        return df
    
    # Convert string sang datetime
    df["CheckinTime"] = pd.to_datetime(df["CheckinTime"], errors="coerce")
    now = datetime.now()

    if period == "week":
        start = now - timedelta(days=now.weekday())
        df = df[df["CheckinTime"] >= start]
    elif period == "month":
        start = datetime(now.year, now.month, 1)
        df = df[df["CheckinTime"] >= start]
    elif period == "year":
        start = datetime(now.year, 1, 1)
        df = df[df["CheckinTime"] >= start]

    return df

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/attendances", methods=["GET"])
def get_attendances():
    try:
        filter_param = request.args.get("filter", "all")

        data = list(collection.find({}, {
            "_id": 0,
            "EmployeeId": 1,
            "EmployeeName": 1,
            "Address": 1,
            "CheckinTime": 1,
            "Status": 1
        }))
        df = pd.DataFrame(data)

        if not df.empty:
            df = filter_dataframe(df, filter_param)
            # Chuyển lại datetime thành chuỗi
            df["CheckinTime"] = df["CheckinTime"].dt.strftime("%Y-%m-%d %H:%M:%S")

        return jsonify(df.to_dict(orient="records")), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/export-excel", methods=["GET"])
def export_to_excel():
    try:
        filter_param = request.args.get("filter", "all")
        data = list(collection.find({}, {
            "_id": 0,
            "EmployeeId": 1,
            "EmployeeName": 1,
            "Address": 1,
            "CheckinTime": 1,
            "Status": 1
        }))
        df = pd.DataFrame(data)
        if not df.empty:
            df = filter_dataframe(df, filter_param)

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
