from flask import Flask, render_template, request, redirect, url_for, flash
from openpyxl import Workbook, load_workbook
import os

app = Flask(__name__)
app.secret_key = "employee_management_secret"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "sheets", "Employees.xlsx")


# -----------------------------
# Create Excel File if Missing
# -----------------------------
def create_excel_file():

    sheets_folder = os.path.join(BASE_DIR, "sheets")

    if not os.path.exists(sheets_folder):
        os.makedirs(sheets_folder)

    if not os.path.exists(EXCEL_FILE):

        wb = Workbook()
        ws = wb.active

        ws.append([
            "Employee ID",
            "Employee Name",
            "Department",
            "Salary"
        ])

        wb.save(EXCEL_FILE)
        wb.close()

    else:

        wb = load_workbook(EXCEL_FILE)
        ws = wb.active

        if ws.max_row == 1 and ws["A1"].value is None:

            ws.append([
                "Employee ID",
                "Employee Name",
                "Department",
                "Salary"
            ])

            wb.save(EXCEL_FILE)

        wb.close()

# -----------------------------
# Get All Employees
# -----------------------------
def get_all_employees():
    create_excel_file()

    wb = load_workbook(EXCEL_FILE)
    ws = wb.active

    employees = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        employees.append({
            "id": row[0],
            "name": row[1],
            "department": row[2],
            "salary": row[3]
        })

    wb.close()
    return employees


# -----------------------------
# Dashboard Statistics
# -----------------------------
def get_dashboard_stats():
    employees = get_all_employees()

    total_employees = len(employees)

    departments = set()
    total_salary = 0

    for emp in employees:
        departments.add(emp["department"])

        try:
            total_salary += float(emp["salary"])
        except:
            pass

    avg_salary = round(
        total_salary / total_employees, 2
    ) if total_employees > 0 else 0

    return {
        "total_employees": total_employees,
        "total_departments": len(departments),
        "avg_salary": avg_salary
    }


# -----------------------------
# Home Page
# -----------------------------
@app.route("/")
def index():
    employees = get_all_employees()
    stats = get_dashboard_stats()

    return render_template(
        "index.html",
        employees=employees,
        stats=stats
    )


# -----------------------------
# Add Employee
# -----------------------------
@app.route("/add", methods=["POST"])
def add_employee():

    employee_id = request.form["employee_id"].strip()
    employee_name = request.form["employee_name"].strip()
    department = request.form["department"].strip()
    salary = request.form["salary"].strip()

    # Validation
    if not employee_id or not employee_name or not department or not salary:
        flash("All fields are required.", "danger")
        return redirect(url_for("index"))

    try:
        float(salary)
    except:
        flash("Salary must be numeric.", "danger")
        return redirect(url_for("index"))

    wb = load_workbook(EXCEL_FILE)
    ws = wb.active

    # Check Duplicate Employee ID
    for row in ws.iter_rows(min_row=2, values_only=True):
        if str(row[0]) == employee_id:
            flash("Employee ID already exists.", "danger")
            wb.close()
            return redirect(url_for("index"))

    ws.append([
        employee_id,
        employee_name,
        department,
        salary
    ])

    print("Saving to:", EXCEL_FILE)

    wb.save(EXCEL_FILE)

    print("Saved Successfully")
    wb.close()

    flash("Employee added successfully.", "success")

    return redirect(url_for("index"))


# -----------------------------
# Edit Employee Page
# -----------------------------
@app.route("/edit/<employee_id>")
def edit_employee(employee_id):

    employees = get_all_employees()

    employee = None

    for emp in employees:
        if str(emp["id"]) == str(employee_id):
            employee = emp
            break

    if employee is None:
        flash("Employee not found.", "danger")
        return redirect(url_for("index"))

    return render_template(
        "edit.html",
        employee=employee
    )


# -----------------------------
# Update Employee
# -----------------------------
@app.route("/update/<employee_id>", methods=["POST"])
def update_employee(employee_id):

    name = request.form["employee_name"].strip()
    department = request.form["department"].strip()
    salary = request.form["salary"].strip()

    if not name or not department or not salary:
        flash("All fields are required.", "danger")
        return redirect(url_for("edit_employee", employee_id=employee_id))

    try:
        float(salary)
    except:
        flash("Salary must be numeric.", "danger")
        return redirect(url_for("edit_employee", employee_id=employee_id))

    wb = load_workbook(EXCEL_FILE)
    ws = wb.active

    found = False

    for row in ws.iter_rows(min_row=2):

        if str(row[0].value) == str(employee_id):

            row[1].value = name
            row[2].value = department
            row[3].value = salary

            found = True
            break

    if found:
        wb.save(EXCEL_FILE)
        flash("Employee updated successfully.", "success")
    else:
        flash("Employee not found.", "danger")

    wb.close()

    return redirect(url_for("index"))


# -----------------------------
# Delete Employee
# -----------------------------
@app.route("/delete/<employee_id>")
def delete_employee(employee_id):

    wb = load_workbook(EXCEL_FILE)
    ws = wb.active

    row_to_delete = None

    for row_num in range(2, ws.max_row + 1):

        if str(ws.cell(row=row_num, column=1).value) == str(employee_id):
            row_to_delete = row_num
            break

    if row_to_delete:
        ws.delete_rows(row_to_delete, 1)

        wb.save(EXCEL_FILE)

        flash(
            "Employee deleted successfully.",
            "success"
        )

    else:
        flash(
            "Employee not found.",
            "danger"
        )

    wb.close()

    return redirect(url_for("index"))


# -----------------------------
# Run App
# -----------------------------
if __name__ == "__main__":
    create_excel_file()
    app.run(debug=True)