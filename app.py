from flask import Flask, render_template, request, jsonify
from flask_cors import CORS
from pymongo import MongoClient
from datetime import datetime, timedelta, timezone
import calendar, re, os

app = Flask(__name__, static_folder="static", template_folder="templates")
CORS(app)

# ---- Timezone VN ----
VN_TZ = timezone(timedelta(hours=7))

# MongoDB config
MONGO_URI = os.getenv(
    "MONGO_URI",
    "mongodb+srv://banhbaobeo2205:lm2hiCLXp6B0D7hq@cluster0.festnla.mongodb.net/?retryWrites=true&w=majority"
)
DB_NAME = os.getenv("DB_NAME", "Sun_Database_1")
client = MongoClient(MONGO_URI)
db = client[DB_NAME]

# -----------------------------
# Hàm build query cho lọc dữ liệu
# -----------------------------
def build_query(filter_type, start_date, end_date, search):
    query = {}
    today = datetime.now(VN_TZ)

    if filter_type == "custom" and start_date and end_date:
        try:
            start = datetime.strptime(start_date, "%Y-%m-%d").replace(tzinfo=VN_TZ)
            end = datetime.strptime(end_date, "%Y-%m-%d").replace(tzinfo=VN_TZ) + timedelta(days=1) - timedelta(seconds=1)
            query["CheckinTime"] = {"$gte": start, "$lte": end}
        except ValueError:
            pass

    elif filter_type == "week":
        start = today - timedelta(days=today.weekday())  # Monday
        start = start.replace(hour=0, minute=0, second=0, microsecond=0)
        end = start + timedelta(days=6, hours=23, minutes=59, seconds=59)
        query["CheckinTime"] = {"$gte": start, "$lte": end}

    elif filter_type == "month":
        start = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        last_day = calendar.monthrange(today.year, today.month)[1]
        end = today.replace(day=last_day, hour=23, minute=59, second=59, microsecond=999999)
        query["CheckinTime"] = {"$gte": start, "$lte": end}

    elif filter_type == "year":
        start = today.replace(month=1, day=1, hour=0, minute=0, second=0, microsecond=0)
        end = today.replace(month=12, day=31, hour=23, minute=59, second=59, microsecond=999999)
        query["CheckinTime"] = {"$gte": start, "$lte": end}

    # Tìm kiếm theo tên NV
    if search:
        query["EmployeeName"] = {"$regex": re.compile(search, re.IGNORECASE)}

    return query


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/api/checkins", methods=["GET"])
def get_checkins():
    filter_type = request.args.get("filter")
    start_date = request.args.get("startDate")
    end_date = request.args.get("endDate")
    search = request.args.get("search")

    query = build_query(filter_type, start_date, end_date, search)

    results = list(db.checkins.find(query).sort("CheckinTime", -1))
    data = []
    for r in results:
        data.append({
            "EmployeeId": r.get("EmployeeId"),
            "EmployeeName": r.get("EmployeeName"),
            "Address": r.get("Address"),
            "Shift": r.get("Shift", "Không xác định"),
            "CheckinTime": r.get("CheckinTime").astimezone(VN_TZ).strftime("%d/%m/%Y %H:%M:%S") if r.get("CheckinTime") else "",
            "Status": r.get("Status"),
            "FaceImage": r.get("FaceImage")
        })

    return jsonify(data)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
