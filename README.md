
# Employee Compensation Forecasting Application

## Objective

Design and build a basic Employee Compensation Forecasting Application for a mid-sized organization. The purpose is to give HR/business stakeholders an interactive way to:

- Analyze current employee compensation
- Apply forecasting adjustments
- Export relevant insights
- Group employees by experience

---

## Tools & Technologies

- **Frontend**: Windows Forms (C#)
- **Backend**: Microsoft SQL Server
- **Data Visualization**: WinForms Chart Control
- **Export**: Excel Interop + CSV via StreamWriter

---

## How to Run

1. Clone or download the repository.
2. Open the solution in **Visual Studio**.
3. Ensure SQL Server is running and database `employee` is properly configured using provided scripts.
4. Press `F5` or click `Start` to run the application.

---

## Features & User Stories Fulfilled

### User Story 1: Filter and Display Active Employees by Role
- Filter by Role and Location
- View Name, Role, Location, Compensation
- Toggle to include/exclude inactive employees
- Average salary dynamically calculated
- Bar chart comparing compensation across locations

### User Story 2: Group by Experience
- One-click chart groups employees by:
  - 0–1 years
  - 1–2 years
  - 2–5 years
  - 5+ years

### User Story 3: Simulate Compensation Increments
- Input custom % increment
- Updated compensation is calculated
- Bonus: Save new compensation to database

### User Story 4: Export Filtered Data
- Export to Excel (.xlsx) using Interop
- Export to CSV (.csv) using built-in file dialog
- Both exports include updated values

---

## Author

Developed as part of a technical case study assignment.

---

