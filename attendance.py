import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
import os
import sys

class StudentRegistry:
    """Handles reading and caching student information from Excel file."""
    
    def __init__(self, filename='studentinfo.xlsx'):
        self.filename = filename
        self.student_cache = {}  # {student_id: rfid}
        self.rfid_cache = {}     # {rfid: student_id}
        self.load_students()
    
    def load_students(self):
        """Load student data from Excel file into memory for fast lookup."""
        try:
            if not os.path.exists(self.filename):
                print(f"Error: {self.filename} not found!")
                return False
            
            wb = openpyxl.load_workbook(self.filename)
            ws = wb.active
            
            # Column A = Student ID, Column K = RFID
            for row in ws.iter_rows(min_row=2, values_only=True):
                student_id = row[0]  # Column A
                rfid = row[10]       # Column K
                
                if student_id and rfid:
                    self.student_cache[str(student_id).strip()] = str(rfid).strip()
                    self.rfid_cache[str(rfid).strip()] = str(student_id).strip()
            
            print(f"✓ Loaded {len(self.student_cache)} students from {self.filename}")
            return True
        except Exception as e:
            print(f"Error loading student data: {e}")
            return False
    
    def get_student_by_id(self, student_id):
        """Get RFID by student ID."""
        return self.student_cache.get(str(student_id).strip())
    
    def get_student_by_rfid(self, rfid):
        """Get student ID by RFID."""
        return self.rfid_cache.get(str(rfid).strip())
    
    def is_valid_student(self, identifier):
        """Check if identifier (student ID or RFID) exists."""
        identifier = str(identifier).strip()
        return identifier in self.student_cache or identifier in self.rfid_cache


class AttendanceSheet:
    """Handles attendance recording and Excel export."""
    
    def __init__(self, filename, staff_name, staff_number, event_name, venue, start_time, end_time):
        self.filename = filename
        self.staff_name = staff_name
        self.staff_number = staff_number
        self.event_name = event_name
        self.venue = venue
        self.start_time = start_time
        self.end_time = end_time
        self.records = []  # List of (timestamp, rfid, student_id) tuples
    
    def add_record(self, rfid, student_id):
        """Add an attendance record."""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.records.append((timestamp, rfid, student_id))
        return timestamp
    
    def save_to_excel(self):
        """Save attendance records to Excel file."""
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Attendance"
            
            # Add header information
            ws['A1'] = "Event Details"
            ws['A1'].font = Font(bold=True, size=12)
            ws['A2'] = f"Staff Name: {self.staff_name}"
            ws['A3'] = f"Staff Number: {self.staff_number}"
            ws['A4'] = f"Event: {self.event_name}"
            ws['A5'] = f"Venue: {self.venue}"
            ws['A6'] = f"Start Time: {self.start_time}"
            ws['A7'] = f"End Time: {self.end_time}"
            ws['A8'] = f"Total Attendees: {len(self.records)}"
            
            # Column headers
            header_row = 10
            ws[f'A{header_row}'] = "Time Scanned"
            ws[f'B{header_row}'] = "RFID"
            ws[f'C{header_row}'] = "Student ID"
            
            # Style headers
            header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF")
            for col in ['A', 'B', 'C']:
                ws[f'{col}{header_row}'].fill = header_fill
                ws[f'{col}{header_row}'].font = header_font
                ws[f'{col}{header_row}'].alignment = Alignment(horizontal="center")
            
            # Add records
            for idx, (timestamp, rfid, student_id) in enumerate(self.records, start=1):
                row = header_row + idx
                ws[f'A{row}'] = timestamp
                ws[f'B{row}'] = rfid
                ws[f'C{row}'] = student_id
            
            # Adjust column widths
            ws.column_dimensions['A'].width = 20
            ws.column_dimensions['B'].width = 25
            ws.column_dimensions['C'].width = 15
            
            wb.save(self.filename)
            print(f"✓ Attendance saved to {self.filename}")
            print(f"✓ Total records: {len(self.records)}")
            return True
        except Exception as e:
            print(f"Error saving attendance file: {e}")
            return False


class AttendanceApp:
    """Main application controller."""
    
    def __init__(self):
        self.student_registry = StudentRegistry()
        self.current_sheet = None
    
    def main_menu(self):
        """Display main menu and handle user choices."""
        while True:
            print("\n" + "="*50)
            print("ATTENDANCE APPLICATION")
            print("="*50)
            print("a. Create new attendance sheet")
            print("b. Load existing attendance sheet")
            print("q. Quit")
            print("="*50)
            
            choice = input("Select option (a/b/q): ").strip().lower()
            
            if choice == 'a':
                self.create_new_sheet()
            elif choice == 'b':
                self.load_existing_sheet()
            elif choice == 'q':
                print("Exiting application. Goodbye!")
                sys.exit(0)
            else:
                print("Invalid choice. Please try again.")
    
    def create_new_sheet(self):
        """Create a new attendance sheet with event details."""
        print("\n" + "="*50)
        print("CREATE NEW ATTENDANCE SHEET")
        print("="*50)
        
        staff_name = input("Staff Name: ").strip()
        staff_number = input("Staff Student Number: ").strip()
        event_name = input("Event Name: ").strip()
        venue = input("Venue: ").strip()
        start_time = input("Event Start Time (HH:MM): ").strip()
        end_time = input("Event End Time (HH:MM): ").strip()
        
        # Generate filename
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"attendance_{event_name.replace(' ', '_')}_{timestamp}.xlsx"
        
        self.current_sheet = AttendanceSheet(
            filename, staff_name, staff_number, event_name, venue, start_time, end_time
        )
        
        print(f"\n✓ Attendance sheet created: {filename}")
        self.record_attendance()
    
    def load_existing_sheet(self):
        """Load an existing attendance sheet from file."""
        print("\nLoading existing attendance sheets...")
        
        attendance_files = [f for f in os.listdir('.') if f.startswith('attendance_') and f.endswith('.xlsx')]
        
        if not attendance_files:
            print("No existing attendance sheets found.")
            return
        
        print("\nAvailable attendance sheets:")
        for idx, filename in enumerate(sorted(attendance_files, reverse=True), 1):
            print(f"{idx}. {filename}")
        
        try:
            choice = int(input("Select file number (or 0 to cancel): ").strip())
            if choice == 0:
                return
            if 1 <= choice <= len(attendance_files):
                filename = sorted(attendance_files, reverse=True)[choice - 1]
                # For now, we'll create a new sheet with loaded filename
                print(f"Loading {filename}...")
                # TODO: Load existing data from file
                print("Append mode not fully implemented yet. Creating new sheet instead.")
                self.create_new_sheet()
            else:
                print("Invalid choice.")
        except ValueError:
            print("Invalid input.")
    
    def record_attendance(self):
        """Record attendance by scanning RFID or entering student ID."""
        if not self.current_sheet:
            print("No active attendance sheet.")
            return
        
        print("\n" + "="*50)
        print("RECORDING ATTENDANCE")
        print("="*50)
        print("Instructions:")
        print("- Scan RFID card or type student ID")
        print("- Type 'save' to save and exit")
        print("- Type 'clear' to clear all records")
        print("="*50)
        
        while True:
            identifier = input("\nScan RFID / Enter Student ID (or 'save'/'clear'): ").strip()
            
            if identifier.lower() == 'save':
                self.current_sheet.save_to_excel()
                self.current_sheet = None
                return
            
            if identifier.lower() == 'clear':
                confirm = input("Clear all records? (y/n): ").strip().lower()
                if confirm == 'y':
                    self.current_sheet.records = []
                    print("✓ All records cleared.")
                continue
            
            if not identifier:
                continue
            
            # Look up student
            student_id = None
            rfid = None
            
            # Check if input is RFID
            if identifier in self.student_registry.rfid_cache:
                rfid = identifier
                student_id = self.student_registry.get_student_by_rfid(rfid)
            # Check if input is student ID
            elif identifier in self.student_registry.student_cache:
                student_id = identifier
                rfid = self.student_registry.get_student_by_id(student_id)
            else:
                print(f"✗ Student not found: {identifier}")
                continue
            
            # Add record
            timestamp = self.current_sheet.add_record(rfid, student_id)
            print(f"✓ {student_id} checked in at {timestamp}")


def main():
    """Entry point."""
    app = AttendanceApp()
    app.main_menu()


if __name__ == "__main__":
    main()
