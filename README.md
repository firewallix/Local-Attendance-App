# Local-Attendance-App (VB6)

A legacy Visual Basic 6 attendance system using Ms. Access and SQL Server backend.

## Features
- Record data in out Employee
- Automatic upload from Access Database to SQL Server (for payroll processing using other app)
- No Reporting
- Automatic download master employee data everytime the app start
- Ignores network connection issues and running 24/7

## Setup Instructions
1. Open `prjAttAccess.vbp` in VB6 IDE
2. Ensure SQL Server is running and database is restored
3. Update connection settings in `absensi.ini`

## Notes
- Tested on Windows 11 with VB6 SP6
- See `docs/setup-guide.md` for full instructions
