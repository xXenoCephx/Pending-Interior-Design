from io import BytesIO
import os
from secrets import choice
from flask import Flask, abort, render_template, session, request, redirect, send_file, url_for
from datetime import datetime, timedelta
import calendar

#Platform Manager Report imports
from Report import Report
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from openpyxl import Workbook

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_KEY', 'dev-secret-key')

DEFAULT_ProjectTags = [
    ['2-Room Flexi', 42], 
    ['3-Room', 87], 
    ['4-Room', 15], 
    ['5-Room', 93], 
    ['3Gen', 61], 
    ['Executive', 28],
    ['HDB', 200],
    ['Condo', 128],
    ['Semi-D', 12],
    ['Bungalow', 7]
]
REPORT_TYPES = [
            "daily report",
            "weekly report",
            "monthly report"
]
DEFAULT_Reports = [
    Report('User Logins', 71, '2026-08-01', 'Daily'),
    Report('User Logins', 91, '2026-08-02', 'Daily'),
    Report('User Logins', 86, '2026-08-03', 'Daily'),
    Report('User Logins', 36, '2026-08-04', 'Daily'),
    Report('User Logins', 70, '2026-08-05', 'Daily'),
    Report('User Logins', 86, '2026-08-06', 'Daily'),
    Report('User Logins', 20, '2026-08-07', 'Daily'),
    Report('User Logins', 13, '2026-08-08', 'Daily'),
    Report('User Logins', 91, '2026-08-09', 'Daily'),
    Report('User Logins', 29, '2026-08-10', 'Daily'),
    Report('User Logins', 35, '2026-08-11', 'Daily'),
    Report('User Logins', 26, '2026-08-12', 'Daily'),
    Report('User Logins', 48, '2026-08-13', 'Daily'),
    Report('User Logins', 4, '2026-08-14', 'Daily'),
    Report('User Logins', 76, '2026-08-15', 'Daily'),
    Report('User Logins', 92, '2026-08-16', 'Daily'),
    Report('User Logins', 9, '2026-08-17', 'Daily'),
    Report('User Logins', 21, '2026-08-18', 'Daily'),
    Report('User Logins', 72, '2026-08-19', 'Daily'),
    Report('User Logins', 91, '2026-08-20', 'Daily'),
    Report('User Logins', 57, '2026-08-21', 'Daily'),
    Report('User Logins', 24, '2026-08-22', 'Daily'),
    Report('User Logins', 78, '2026-08-23', 'Daily'),
    Report('User Logins', 95, '2026-08-24', 'Daily'),
    Report('User Logins', 70, '2026-08-25', 'Daily'),
    Report('User Logins', 91, '2026-08-26', 'Daily'),
    Report('User Logins', 37, '2026-08-27', 'Daily'),
    Report('User Logins', 70, '2026-08-28', 'Daily'),
    Report('User Logins', 45, '2026-08-29', 'Daily'),
    Report('User Logins', 32, '2026-08-30', 'Daily'),
    Report('User Logins', 96, '2026-08-31', 'Daily'),

    Report('User Logins', 31, '2026-09-01', 'Daily'),
    Report('User Logins', 18, '2026-09-02', 'Daily'),
    Report('User Logins', 9, '2026-09-03', 'Daily'),
    Report('User Logins', 41, '2026-09-04', 'Daily'),
    Report('User Logins', 29, '2026-09-05', 'Daily'),
    Report('User Logins', 97, '2026-09-06', 'Daily'),
    Report('User Logins', 69, '2026-09-07', 'Daily'),
    Report('User Logins', 29, '2026-09-08', 'Daily'),
    Report('User Logins', 93, '2026-09-09', 'Daily'),
    Report('User Logins', 85, '2026-09-10', 'Daily'),
    Report('User Logins', 5, '2026-09-11', 'Daily'),
    Report('User Logins', 54, '2026-09-12', 'Daily'),
    Report('User Logins', 85, '2026-09-13', 'Daily'),
    Report('User Logins', 71, '2026-09-14', 'Daily'),
    Report('User Logins', 52, '2026-09-15', 'Daily'),
    Report('User Logins', 6, '2026-09-16', 'Daily'),
    Report('User Logins', 21, '2026-09-17', 'Daily'),
    Report('User Logins', 10, '2026-09-18', 'Daily'),
    Report('User Logins', 63, '2026-09-19', 'Daily'),
    Report('User Logins', 35, '2026-09-20', 'Daily'),
    Report('User Logins', 2, '2026-09-21', 'Daily'),
    Report('User Logins', 46, '2026-09-22', 'Daily'),
    Report('User Logins', 61, '2026-09-23', 'Daily'),
    Report('User Logins', 49, '2026-09-24', 'Daily'),
    Report('User Logins', 56, '2026-09-25', 'Daily'),
    Report('User Logins', 40, '2026-09-26', 'Daily'),
    Report('User Logins', 13, '2026-09-27', 'Daily'),
    Report('User Logins', 49, '2026-09-28', 'Daily'),
    Report('User Logins', 78, '2026-09-29', 'Daily'),
    Report('User Logins', 63, '2026-09-30', 'Daily'),

    Report('User Logins', 6, '2026-10-01', 'Daily'),
    Report('User Logins', 21, '2026-10-02', 'Daily'),
    Report('User Logins', 45, '2026-10-03', 'Daily'),
    Report('User Logins', 63, '2026-10-04', 'Daily'),
    Report('User Logins', 30, '2026-10-05', 'Daily'),
    Report('User Logins', 71, '2026-10-06', 'Daily'),
    Report('User Logins', 98, '2026-10-07', 'Daily'),
    Report('User Logins', 73, '2026-10-08', 'Daily'),
    Report('User Logins', 12, '2026-10-09', 'Daily'),
    Report('User Logins', 80, '2026-10-10', 'Daily'),
    Report('User Logins', 5, '2026-10-11', 'Daily'),
    Report('User Logins', 17, '2026-10-12', 'Daily'),
    Report('User Logins', 13, '2026-10-13', 'Daily'),
    Report('User Logins', 33, '2026-10-14', 'Daily'),
    Report('User Logins', 27, '2026-10-15', 'Daily'),
    Report('User Logins', 89, '2026-10-16', 'Daily'),
    Report('User Logins', 24, '2026-10-17', 'Daily'),
    Report('User Logins', 56, '2026-10-18', 'Daily'),
    Report('User Logins', 48, '2026-10-19', 'Daily'),
    Report('User Logins', 3, '2026-10-20', 'Daily'),
    Report('User Logins', 97, '2026-10-21', 'Daily'),
    Report('User Logins', 44, '2026-10-22', 'Daily'),
    Report('User Logins', 23, '2026-10-23', 'Daily'),
    Report('User Logins', 37, '2026-10-24', 'Daily'),
    Report('User Logins', 100, '2026-10-25', 'Daily'),
    Report('User Logins', 65, '2026-10-26', 'Daily'),
    Report('User Logins', 68, '2026-10-27', 'Daily'),
    Report('User Logins', 40, '2026-10-28', 'Daily'),
    Report('User Logins', 65, '2026-10-29', 'Daily'),
    Report('User Logins', 54, '2026-10-30', 'Daily'),
    Report('User Logins', 59, '2026-10-31', 'Daily'),

    Report('User Logins', 30, '2026-11-01', 'Daily'),
    Report('User Logins', 61, '2026-11-02', 'Daily'),
    Report('User Logins', 48, '2026-11-03', 'Daily'),
    Report('User Logins', 81, '2026-11-04', 'Daily'),
    Report('User Logins', 24, '2026-11-05', 'Daily'),
    Report('User Logins', 89, '2026-11-06', 'Daily'),
    Report('User Logins', 65, '2026-11-07', 'Daily'),
    Report('User Logins', 17, '2026-11-08', 'Daily'),
    Report('User Logins', 62, '2026-11-09', 'Daily'),
    Report('User Logins', 52, '2026-11-10', 'Daily'),
    Report('User Logins', 48, '2026-11-11', 'Daily')
]


@app.route('/')
def home():    
    return render_template("sites/Main.html")

@app.route('/login')
def login():
    return render_template("sites/Login.html")

@app.route('/loginCheck', methods=["POST"])
def loginCheck():
    email = request.form.get("email")
    password = request.form.get("password")
    if email == "Testing@email.com" and password == "123456":
        session['user'] = "Testing"
    elif email == "pm@email.com" and password == "123456":
        session['user'] = "PlatformManager"
        return redirect('/PlatformManagers')

    return redirect('/')

@app.route('/logout')
def logout():
    session.pop('user', None)
    return redirect("/")

@app.route('/Designers')
def designers():
    return render_template("sites/Designers.html")

@app.route("/SignUp")
def signup():
    return render_template("sites/SignUp.html")

@app.route('/Projects')
def projects():
    return render_template('sites/Main.html')

@app.route('/PlatformManagers')
def platform_managers():
    return render_template("sites/PlatformManagers.html")

@app.route('/GenerateReport')
def generate_report():
    
    today_date = datetime.today().strftime('%Y-%m-%d')
    return render_template("sites/GenerateReport.html",report_types=REPORT_TYPES, default_date=today_date)

@app.route('/GenerateReport/create', methods=["POST"])
def create_report():
    to_date = lambda s: datetime.strptime(s, '%Y-%m-%d').date()
    raw_reports = [r for r in DEFAULT_Reports if r.report_type not in REPORT_TYPES]
    dates = [to_date(r.report_date) for r in raw_reports]
    first_available, last_available = min(dates), max(dates)

    report_type = request.form.get("reportType", "")
    session["generate_form"] = {"reportType": report_type, "startDate": request.form.get("startDate", "")}

    start = to_date(request.form.get("startDate", ""))

    if report_type == "weekly report":
        requested_end = start + timedelta(days=6)
    elif report_type == "monthly report":
        year, month = (start.year + 1, 1) if start.month == 12 else (start.year, start.month + 1)
        day = min(start.day, calendar.monthrange(year, month)[1])
        requested_end = start.replace(year=year, month=month, day=day) - timedelta(days=1)
    else:  # daily report
        requested_end = start

    end = min(requested_end, last_available)  # cut off at the last available day
    total = sum(r.report_loginvalue for r in raw_reports if start <= to_date(r.report_date) <= end)

    # Create the Report object and store it (replaces an identical earlier one)
    new_report = Report("User Logins", total, start.isoformat(), report_type)
    DEFAULT_Reports[:] = [r for r in DEFAULT_Reports
                            if not (r.report_type == new_report.report_type
                                    and r.report_date == new_report.report_date)]
    DEFAULT_Reports.append(new_report)

    session["generated"] = {
        "type": report_type,
        "start": start.isoformat(),
        "end": end.isoformat(),
        "requested_end": requested_end.isoformat(),
    }

    return redirect(url_for('generate_report'))

@app.route('/ExportReport')
def export_report():
    return render_template("sites/ExportReport.html", reports=DEFAULT_Reports)

@app.route('/ExportReport/pdf', methods=["POST"])
def export_report_pdf():
    choice = request.form.get("report_index", "all")
    selected = [DEFAULT_Reports[int(choice)]]
    filename = f"report_{selected[0].report_date}.pdf"

    # Building PDF 
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, title="Report")
    styles = getSampleStyleSheet()

    rows = [["Date", "Type", "Value"]]
    rows += [[r.report_date, r.report_logins, str(r.report_loginvalue)] for r in selected]
    if len(selected) > 1:
        rows.append(["Total", "", str(sum(r.report_loginvalue for r in selected))])

    table = Table(rows, colWidths=[150, 200, 100], repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (2, 0), (2, -1), "RIGHT"),
    ]))

    doc.build([Paragraph("User Login Report", styles["Title"]), Spacer(1, 12), table])
    buffer.seek(0)
    return send_file(buffer, mimetype="application/pdf",
                     as_attachment=True, download_name=filename)

@app.route('/ExportReport/excel', methods=["POST"])
def export_report_excel():
    choice = request.form.get("report_index", "all")
    selected = [DEFAULT_Reports[int(choice)]]
    suffix = selected[0].report_date
    filename = f"report_{suffix}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = f"Report - {suffix}"
    # Title
    ws.append(["Date", "Type", "Value"])
    for r in selected:
        ws.append([r.report_date, r.report_logins, r.report_loginvalue])

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return send_file(
        buffer,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"report_{suffix}.xlsx",
    )

@app.route('/About')
def about():
    return render_template('sites/Main.html')


if __name__ == "__main__":
    app.run(debug=True, port=5001)
