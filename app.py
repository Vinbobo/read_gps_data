from flask import Flask, render_template, jsonify
from pymongo import MongoClient
from flask_cors import CORS
import os

app = Flask(__name__, template_folder="templates")
CORS(app)

# 🔹 MongoDB Atlas connection string (đặt trong Render -> Environment Variables)
MONGO_URI = os.getenv(
    "MONGO_URI",
    "mongodb+srv://banhbaobeo2205:lm2hiCLXp6B0D7hq@cluster0.festnla.mongodb.net/?retryWrites=true&w=majority"
)
DB_NAME = os.getenv("DB_NAME", "Sun_Database_1")

client = MongoClient(MONGO_URI)
db = client[DB_NAME]
collection = db["HR_GPS_Attendance"]

# 🔹 Trang chính
@app.route("/")
def index():
    return render_template("index.html")

# 🔹 REST API trả dữ liệu
@app.route("/api/attendances", methods=["GET"])
def get_attendances():
    try:
        # chỉ lấy các field cần hiển thị
        data = list(collection.find({}, {
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

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
