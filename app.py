from flask import Flask, render_template, request, redirect, session
import pandas as pd
from config import FILE_PATH, SECRET_KEY

app = Flask(__name__)
app.secret_key = SECRET_KEY


# ================= LOAD DATA =================
def load_data():
    data = {
        "login": pd.read_excel(FILE_PATH, sheet_name="login_users"),
        "students": pd.read_excel(FILE_PATH, sheet_name="students"),
        "attendance": pd.read_excel(FILE_PATH, sheet_name="attendance"),
        "marks": pd.read_excel(FILE_PATH, sheet_name="marks"),
        "fees": pd.read_excel(FILE_PATH, sheet_name="fees"),
        "assignments": pd.read_excel(FILE_PATH, sheet_name="assignments"),
        "notifications": pd.read_excel(FILE_PATH, sheet_name="notifications"),
        "exams": pd.read_excel(FILE_PATH, sheet_name="exam_schedule"),
        "library": pd.read_excel(FILE_PATH, sheet_name="library"),
        "placement": pd.read_excel(FILE_PATH, sheet_name="placements"),
        "materials": pd.read_excel(FILE_PATH, sheet_name="course_material"),
        "timetable": pd.read_excel(FILE_PATH, sheet_name="timetable")
    }

    for key in data:
        data[key].columns = data[key].columns.str.strip()

    return data


# ================= LOGIN =================
@app.route("/")
def home():
    return render_template("login.html")


@app.route("/login", methods=["POST"])
def login():
    data = load_data()
    login_df = data["login"]

    user = login_df[
        (login_df["username"] == request.form["username"]) &
        (login_df["password"] == request.form["password"])
    ]

    if not user.empty:
        session["role"] = user.iloc[0]["role"]
        session["student_id"] = user.iloc[0]["student_id"]
        session.pop("selected_student", None)

        if session["role"] == "admin":
            return redirect("/admin")

        return redirect("/dashboard")

    return "Invalid Credentials"


# ================= ADMIN =================
@app.route("/admin")
def admin():
    if session.get("role") != "admin":
        return redirect("/")

    data = load_data()
    return render_template("admin_students.html",
                           students=data["students"].to_dict(orient="records"))


@app.route("/admin/select/<int:sid>")
def admin_select(sid):
    if session.get("role") != "admin":
        return redirect("/")

    session["selected_student"] = sid
    return redirect("/dashboard")


@app.route("/add_student")
def add_student():
    if session.get("role") != "admin":
        return redirect("/")
    return render_template("add_student.html")


@app.route("/save_student", methods=["POST"])
def save_student():
    if session.get("role") != "admin":
        return redirect("/")

    data = load_data()
    students_df = data["students"]

    new_id = students_df["student_id"].max() + 1

    new_student = {
        "student_id": new_id,
        "student_name": request.form["student_name"],
        "department": request.form["department"],
        "year": request.form["year"],
        "email": request.form["email"],
        "phone": request.form["phone"]
    }

    students_df = pd.concat([students_df, pd.DataFrame([new_student])],
                            ignore_index=True)

    with pd.ExcelWriter(FILE_PATH, engine="openpyxl",
                        mode="a", if_sheet_exists="replace") as writer:
        students_df.to_excel(writer, sheet_name="students", index=False)

    return redirect("/admin")


# ================= HELPER =================
def current_student():
    if session.get("role") == "admin":
        return session.get("selected_student")
    return session.get("student_id")


def require_student():
    sid = current_student()
    if sid is None:
        return None
    return sid


# ================= DASHBOARD =================
@app.route("/dashboard")
def dashboard():
    sid = require_student()
    if sid is None:
        return redirect("/admin")

    data = load_data()
    student = data["students"][data["students"]["student_id"] == sid].iloc[0]

    return render_template("dashboard_home.html", student=student)


# ================= PROFILE =================
@app.route("/profile")
def profile():
    sid = require_student()
    if sid is None:
        return redirect("/admin")

    data = load_data()
    student = data["students"][data["students"]["student_id"] == sid].iloc[0]

    return render_template("profile.html", student=student)


# ================= TIMETABLE =================
@app.route("/timetable")
def timetable():
    sid = require_student()
    if sid is None:
        return redirect("/admin")

    data = load_data()
    student = data["students"][data["students"]["student_id"] == sid].iloc[0]
    dept = student["department"]

    timetable_df = data["timetable"]
    dept_column = "Dept" if "Dept" in timetable_df.columns else "Department"

    timetable_df = timetable_df[timetable_df[dept_column] == dept]

    selected_day = request.args.get("day")
    if selected_day:
        timetable_df = timetable_df[timetable_df["Day"] == selected_day]

    days = timetable_df["Day"].unique()

    return render_template("timetable.html",
                           timetable=timetable_df.to_dict(orient="records"),
                           days=days,
                           selected_day=selected_day)


# ================= ATTENDANCE =================
@app.route("/attendance")
def attendance():
    sid = require_student()
    if sid is None:
        return redirect("/admin")

    data = load_data()
    records = data["attendance"][data["attendance"]["student_id"] == sid]

    return render_template("attendance.html",
                           attendance=records.to_dict(orient="records"),
                           role=session.get("role"))


@app.route("/attendance/edit/<int:sid>/<date>")
def edit_attendance(sid, date):
    if session.get("role") != "admin":
        return redirect("/attendance")

    data = load_data()
    df = data["attendance"]

    record = df[(df["student_id"] == sid) & (df["date"] == date)]

    return render_template("attendance_edit.html",
                       record=record.iloc[0])



@app.route("/attendance/update", methods=["POST"])
def update_attendance():
    if session.get("role") != "admin":
        return redirect("/attendance")

    data = load_data()
    df = data["attendance"]

    sid = int(request.form["student_id"])
    date = request.form["date"]

    df.loc[
        (df["student_id"] == sid) & (df["date"] == date),
        ["period1", "period2", "period3", "period4"]
    ] = [
        request.form["period1"],
        request.form["period2"],
        request.form["period3"],
        request.form["period4"]
    ]

    with pd.ExcelWriter(FILE_PATH, engine="openpyxl",
                        mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="attendance", index=False)

    return redirect("/attendance")


# ================= MARKS =================

@app.route("/marks")
def marks():
    sid = require_student()
    if sid is None:
        return redirect("/admin")

    data = load_data()
    df = data["marks"]

    records = df[df["student_id"] == sid]

    return render_template("marks.html",
                           marks=records.to_dict(orient="records"),
                           role=session.get("role"))


@app.route("/marks/edit/<int:sid>")
def edit_marks(sid):
    if session.get("role") != "admin":
        return redirect("/marks")

    data = load_data()
    df = data["marks"]

    record = df[df["student_id"] == sid]

    if record.empty:
        return "Marks record not found"

    return render_template("marks_edit.html",
                           record=record.iloc[0])



@app.route("/marks/update", methods=["POST"])
def update_marks():
    if session.get("role") != "admin":
        return redirect("/marks")

    data = load_data()
    df = data["marks"]

    sid = int(request.form["student_id"])
    internal = int(request.form["internal"])
    external = int(request.form["external"])

    total = internal + external
    result = "Pass" if total >= 40 else "Fail"

    df.loc[df["student_id"] == sid,
           ["internal", "external", "total", "result"]] = \
        [internal, external, total, result]

    with pd.ExcelWriter(FILE_PATH, engine="openpyxl",
                        mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="marks", index=False)

    return redirect("/marks")

# ================= FEES =================
@app.route("/fees")
def fees():
    sid = require_student()
    if sid is None:
        return redirect("/admin")

    data = load_data()
    records = data["fees"][data["fees"]["student_id"] == sid]

    return render_template("fees.html",
                           fees=records.to_dict(orient="records"),
                           role=session.get("role"))


@app.route("/fees/edit/<int:sid>")
def edit_fees(sid):
    if session.get("role") != "admin":
        return redirect("/fees")

    data = load_data()
    record = data["fees"][data["fees"]["student_id"] == sid]

    return render_template("fees_edit.html",
                       record=record.iloc[0])



@app.route("/fees/update", methods=["POST"])
def update_fees():
    if session.get("role") != "admin":
        return redirect("/fees")

    data = load_data()
    df = data["fees"]

    sid = int(request.form["student_id"])
    paid = int(request.form["paid_fee"])

    total = df.loc[df["student_id"] == sid, "total_fee"].values[0]
    balance = total - paid
    status = "Paid" if balance == 0 else "Pending"

    df.loc[df["student_id"] == sid,
           ["paid_fee", "balance", "payment_status"]] = \
        [paid, balance, status]

    with pd.ExcelWriter(FILE_PATH, engine="openpyxl",
                        mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="fees", index=False)

    return redirect("/fees")


# ================= OTHER MODULES =================
@app.route("/assignments")
def assignments():
    data = load_data()
    return render_template("assignments.html",
                           assignments=data["assignments"].to_dict(orient="records"))


@app.route("/notifications")
def notifications():
    data = load_data()
    return render_template("notifications.html",
                           notifications=data["notifications"].to_dict(orient="records"))


@app.route("/exams")
def exams():
    data = load_data()
    return render_template("exams.html",
                           exams=data["exams"].to_dict(orient="records"))


@app.route("/library")
def library():
    sid = require_student()
    if sid is None:
        return redirect("/admin")

    data = load_data()
    records = data["library"][data["library"]["student_id"] == sid]

    return render_template("library.html",
                           library=records.to_dict(orient="records"))


@app.route("/placement")
def placement():
    sid = require_student()
    if sid is None:
        return redirect("/admin")

    data = load_data()
    records = data["placement"][data["placement"]["student_id"] == sid]

    return render_template("placement.html",
                           placement=records.to_dict(orient="records"))


@app.route("/materials")
def materials():
    data = load_data()
    return render_template("materials.html",
                           materials=data["materials"].to_dict(orient="records"))


@app.route("/attendance_summary")
def attendance_summary():
    if session.get("role") != "admin":
        return redirect("/")

    data = load_data()
    attendance_df = data["attendance"]

    # Identify period columns automatically
    period_columns = [col for col in attendance_df.columns if "period" in col.lower()]

    if not period_columns:
        return "No period columns found in attendance sheet."

    summary = []

    total_present = 0
    total_absent = 0

    # Group by student
    grouped = attendance_df.groupby("student_id")

    for sid, group in grouped:
        student_name = group.iloc[0]["student_name"]

        total_classes = 0
        present_count = 0

        for col in period_columns:
            total_classes += group[col].count()
            present_count += (group[col] == "P").sum()

        absent_count = total_classes - present_count
        percentage = round((present_count / total_classes) * 100, 2) if total_classes > 0 else 0

        total_present += present_count
        total_absent += absent_count

        summary.append({
            "student_id": sid,
            "student_name": student_name,
            "total_classes": total_classes,
            "present": present_count,
            "absent": absent_count,
            "percentage": percentage
        })

    overall_total = total_present + total_absent
    overall_percentage = round((total_present / overall_total) * 100, 2) if overall_total > 0 else 0

    return render_template(
        "attendance_summary.html",
        summary=summary,
        overall_percentage=overall_percentage
    )


# ================= LOGOUT =================
@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
