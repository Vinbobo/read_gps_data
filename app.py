from flask import Flask, render_template, jsonify, send_file, request
from pymongo import MongoClient
from flask_cors import CORS
import os
import pandas as pd
from io import BytesIO
from datetime import datetime

app = Flask(__name__, template_folder="templates")
CORS(app)

# 🔹 MongoDB Atlas connection
MONGO_URI = os.getenv("MONGO_URI", "mongodb+srv://....")
DB_NAME = os.getenv("DB_NAME", "Sun_Database_1")

client = MongoClient(MONGO_URI)
db = client[DB_NAME]
collection = db["checkins"]

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/attendances", methods=["GET"])
def get_attendances():
    try:
        filter_type = request.args.get("filter", "all")
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")

        query = {}
        if start_date and end_date:
            query["CheckinTime"] = {
                "$gte": start_date,
                "$lte": end_date
            }

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

@app.route("/api/export-excel", methods=["GET"])
def export_to_excel():
    try:
        start_date = request.args.get("startDate")
        end_date = request.args.get("endDate")

        query = {}
        if start_date and end_date:
            query["CheckinTime"] = {
                "$gte": start_date,
                "$lte": end_date
            }

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

        # 👉 Gợi ý tên file
        if start_date and end_date:
            filename = f"attendance_{start_date}_to_{end_date}.xlsx"
        else:
            today = datetime.now().strftime("%Y%m%d")
            filename = f"attendance_{today}.xlsx"

        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
