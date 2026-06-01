````markdown name=README.md url=https://github.com/asm0git/attendance-app/blob/main/README.md
# Attendance System - RFID Scanner with Modern GUI

A fast, responsive, and user-friendly Python attendance tracking application using RFID scanner input with a modern GUI interface.

## ✨ Features

- **🎯 Modern GUI Interface** - Clean, intuitive design with minimal clicks
- **⚡ Lightning-Fast Performance** - Optimized for high-volume scanning
- **🔄 RFID Scanner Support** - Instant recognition of RFID card scans
- **📝 Manual Entry** - Type student ID for quick entry
- **💾 Auto-Save** - Prevents duplicate entries automatically
- **📊 Excel Export** - Professional formatted attendance files
- **🔌 Load Existing Sheets** - Resume taking attendance from previous sessions

## 🚀 Quick Start

### Installation

```bash
# Clone or download the project
cd attendance-app

# Install dependencies
pip install -r requirements.txt
```

### Setup

1. **Prepare Student Data** (Required)
   - Create a file named `studentinfo.xlsx` in the same folder as `attendance.py`
   - **Column A**: Student ID (e.g., S001, S002, S003)
   - **Column B**: Student Name (optional, for display)
   - **Column K**: RFID (e.g., 1234567890, 0987654321)

   Example structure:
   ```
   Student ID | First Name | ... (9 more columns) | RFID
   S001       | John       | ... |                  | 1234567890
   S002       | Jane       | ... |                  | 0987654321
   ```

2. **Run the Application**
   ```bash
   python attendance.py
   ```

## 📖 Usage Guide

### Main Menu
- **CREATE NEW ATTENDANCE** - Start a new attendance session
- **LOAD EXISTING ATTENDANCE** - Continue from a previous session
- **EXIT** - Close the application

### Creating New Attendance Sheet

1. Click "CREATE NEW ATTENDANCE"
2. Fill in the form:
   - Staff Name
   - Staff Student Number
   - Event Name
   - Venue
   - Start Time (HH:MM format)
   - End Time (HH:MM format)
3. Click "START SCANNING"

### Scanning Students

The scanning interface is optimized for speed:
- **Input field is auto-focused** - Ready for RFID scanner input immediately
- **Press Enter or scan RFID** - Instant recognition
- **Type student ID** - Alternative to RFID (e.g., S001)
- **Real-time feedback** - See success/error messages instantly
- **Duplicate prevention** - Same student can only be marked once
- **Live count** - Shows total students scanned

### Controls During Scanning

- **💾 SAVE & FINISH** - Export to Excel and return to menu
- **🔄 CLEAR ALL** - Delete all scanned records (with confirmation)
- **← BACK** - Return to menu (without saving)

## 📁 Output Files

Attendance sheets are automatically saved as:
```
attendance_EventName_YYYYMMDD_HHMMSS.xlsx
```

### Excel File Format

| Column | Content |
|--------|---------|
| A | Time Scanned (timestamp) |
| B | RFID |
| C | Student ID |

Header section includes:
- Staff information
- Event details
- Total attendees count

## ⚙️ System Requirements

- Python 3.7+
- openpyxl library (handles Excel files)
- Tkinter (included with Python)

## 🎮 Performance Tips

- **Minimize interventions** - Input field auto-focuses after each scan
- **Color feedback** - Green = success, Red = error/duplicate
- **Scrollable list** - Automatically scrolls to show latest entry
- **Duplicate handling** - Automatically prevents re-scanning same student

## 🔧 Troubleshooting

### "studentinfo.xlsx not found"
- Create the file in the same directory as `attendance.py`
- Ensure Column A has Student IDs and Column K has RFIDs

### RFID not recognized
- Check that the RFID number matches exactly in Column K
- Verify no extra spaces in the Excel file

### GUI scaling issues
- Adjust window size by dragging edges
- Font sizes are scaled automatically

## 📝 License

Free to use and modify for educational purposes.

## 🚀 Future Enhancements

- Photo capture integration
- Network-based student database
- Real-time attendance tracking dashboard
- Mobile app support
- QR code scanning

---

**Created for fast, efficient attendance tracking with minimal human intervention!**
````
