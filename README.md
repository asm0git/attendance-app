````markdown name=README.md url=https://github.com/asm0git/attendance-app/blob/main/README.md
# Attendance System - RFID Scanner with Modern GUI

A fast, responsive, and user-friendly Python attendance tracking application using RFID scanner input with comprehensive logging and a modern GUI interface.

## ✨ Features

- **🎯 Modern GUI Interface** - Clean, intuitive design with minimal clicks
- **⚡ Lightning-Fast Performance** - Optimized for high-volume scanning
- **🔄 RFID Scanner Support** - Instant recognition of RFID card scans
- **📝 Manual Entry** - Type student ID for quick entry
- **💾 Auto-Save** - Prevents duplicate entries automatically
- **📊 Excel Export** - Professional formatted attendance files with multiple sheets
- **📋 Complete Logging** - ALL scanned IDs are logged, whether valid or unknown
- **🔌 Load Existing Sheets** - Resume taking attendance from previous sessions
- **✓ Zero Data Loss** - Every input is recorded and preserved

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
- **ALL inputs are logged** - Even unknown IDs are recorded
- **Duplicate prevention** - Same valid student can only be marked once
- **Live count** - Shows valid attendees + total log entries

### Controls During Scanning

- **💾 SAVE & FINISH** - Export to Excel and return to menu
- **🔄 CLEAR ALL** - Delete all scanned records (with confirmation)
- **← BACK** - Return to menu (without saving)

## 📁 Output Files

Attendance sheets are automatically saved as:
```
attendance_EventName_YYYYMMDD_HHMMSS.xlsx
```

### Excel File Structure (3 Sheets)

#### **Sheet 1: Summary**
- Event information
- Staff details
- Statistics including:
  - Total inputs scanned
  - Valid attendees count
  - Unknown/invalid IDs count

#### **Sheet 2: Attendance** (Valid entries only)
| Column | Content |
|--------|---------|
| A | Time Scanned (timestamp) |
| B | RFID |
| C | Student ID |

**Header section includes:**
- Staff information
- Event details
- Total valid attendees count

#### **Sheet 3: Complete Log** (ALL inputs)
| Column | Content |
|--------|---------|
| A | Timestamp |
| B | Input Scanned (what was entered) |
| C | RFID |
| D | Student ID |
| E | Status |

**Status values:**
- `FOUND_BY_RFID` - Scanned/entered RFID matched a student (green)
- `FOUND_BY_ID` - Scanned/entered student ID matched a record (green)
- `UNKNOWN` - ID/RFID not found in studentinfo.xlsx (red highlight)

**Key Features:**
- ✅ **Zero data loss** - Every input is recorded
- ✅ **Complete audit trail** - All scans logged with timestamps
- ✅ **Color-coded status** - Easy identification of valid vs unknown entries
- ✅ **Traceable history** - Know exactly what was scanned when

## ⚙️ System Requirements

- Python 3.7+
- openpyxl library (handles Excel files)
- Tkinter (included with Python)

## 🎮 Performance & Usage Tips

### Minimize Interventions
- Input auto-focuses after each scan
- Press Enter = instant processing
- RFID scanner = automatic detection
- Manual typing = student ID lookup
- Color-coded feedback (no need to read text)
- One button to save & finish

### Data Integrity
- **All inputs saved** - No data is lost
- **Complete logging** - Every action is timestamped
- **Audit trail** - Review all scans in Complete Log sheet
- **Unknown tracking** - Identify invalid entries for follow-up

## 🔧 Troubleshooting

### "studentinfo.xlsx not found"
- Create the file in the same directory as `attendance.py`
- Ensure Column A has Student IDs and Column K has RFIDs

### RFID not recognized
- Check that the RFID number matches exactly in Column K
- Unknown IDs are logged in the Complete Log sheet for review
- Verify no extra spaces in the Excel file

### GUI scaling issues
- Adjust window size by dragging edges
- Font sizes are scaled automatically

## 📝 Data Backup & Recovery

All attendance files are saved locally and can be:
- **Reopened** - Load existing sheets to add more entries
- **Analyzed** - Review the Complete Log for unknown entries
- **Audited** - Check timestamps and status of all scans
- **Exported** - Use Excel to further analyze or print

## 🚀 Future Enhancements

- Photo capture integration
- Network-based student database
- Real-time attendance tracking dashboard
- Mobile app support
- QR code scanning
- Automatic email reports

---

**Created for fast, efficient, accurate attendance tracking with complete data logging and zero data loss!**
````
