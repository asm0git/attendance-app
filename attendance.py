import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
import os
import threading
from collections import OrderedDict

class StudentRegistry:
    """Handles reading and caching student information from Excel file."""
    
    def __init__(self, filename='studentinfo.xlsx'):
        self.filename = filename
        self.student_cache = {}  # {student_id: rfid}
        self.rfid_cache = {}     # {rfid: student_id}
        self.student_names = {}  # {student_id: name} for display
        self.load_students()
    
    def load_students(self):
        """Load student data from Excel file into memory for fast lookup."""
        try:
            if not os.path.exists(self.filename):
                return False, f"Error: {self.filename} not found!"
            
            wb = openpyxl.load_workbook(self.filename)
            ws = wb.active
            
            # Column A = Student ID, Column K = RFID, Column B = Name (optional)
            for row in ws.iter_rows(min_row=2, values_only=True):
                student_id = row[0]  # Column A
                name = row[1] if row[1] else ""  # Column B (optional)
                rfid = row[10]       # Column K
                
                if student_id and rfid:
                    self.student_cache[str(student_id).strip()] = str(rfid).strip()
                    self.rfid_cache[str(rfid).strip()] = str(student_id).strip()
                    self.student_names[str(student_id).strip()] = str(name).strip()
            
            return True, f"Loaded {len(self.student_cache)} students"
        except Exception as e:
            return False, f"Error loading student data: {e}"
    
    def get_student_by_id(self, student_id):
        """Get RFID by student ID."""
        return self.student_cache.get(str(student_id).strip())
    
    def get_student_by_rfid(self, rfid):
        """Get student ID by RFID."""
        return self.rfid_cache.get(str(rfid).strip())
    
    def get_student_name(self, student_id):
        """Get student name if available."""
        return self.student_names.get(str(student_id).strip(), "")
    
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
        self.records = OrderedDict()  # {student_id: (timestamp, rfid)} to prevent duplicates
    
    def add_record(self, rfid, student_id):
        """Add an attendance record (prevents duplicates)."""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        if student_id not in self.records:
            self.records[student_id] = (timestamp, rfid)
            return timestamp, True
        return self.records[student_id][0], False  # Return existing time, False for duplicate
    
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
            for idx, (student_id, (timestamp, rfid)) in enumerate(self.records.items(), start=1):
                row = header_row + idx
                ws[f'A{row}'] = timestamp
                ws[f'B{row}'] = rfid
                ws[f'C{row}'] = student_id
            
            # Adjust column widths
            ws.column_dimensions['A'].width = 20
            ws.column_dimensions['B'].width = 25
            ws.column_dimensions['C'].width = 15
            
            wb.save(self.filename)
            return True, f"Saved {len(self.records)} records to {self.filename}"
        except Exception as e:
            return False, f"Error saving attendance file: {e}"


class AttendanceUI:
    """Modern, fast, responsive UI with minimal human intervention."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Attendance System - RFID Scanner")
        self.root.geometry("1000x700")
        self.root.configure(bg='#f0f0f0')
        
        self.student_registry = StudentRegistry()
        self.current_sheet = None
        self.scanned_students = []
        
        self.setup_styles()
        self.show_menu()
    
    def setup_styles(self):
        """Setup TTK styles for modern look."""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Colors
        style.configure('Title.TLabel', font=('Helvetica', 24, 'bold'), background='#f0f0f0', foreground='#1a5490')
        style.configure('Header.TLabel', font=('Helvetica', 14, 'bold'), background='#f0f0f0')
        style.configure('Normal.TLabel', font=('Helvetica', 11), background='#f0f0f0')
        style.configure('Success.TLabel', font=('Helvetica', 12, 'bold'), foreground='#28a745', background='#f0f0f0')
        style.configure('Error.TLabel', font=('Helvetica', 12, 'bold'), foreground='#dc3545', background='#f0f0f0')
        style.configure('Info.TLabel', font=('Helvetica', 10), background='#f0f0f0', foreground='#666')
        
        style.configure('Action.TButton', font=('Helvetica', 12, 'bold'), padding=10)
        style.configure('TEntry', font=('Helvetica', 11), padding=5)
    
    def clear_window(self):
        """Clear all widgets from window."""
        for widget in self.root.winfo_children():
            widget.destroy()
    
    def show_menu(self):
        """Show main menu."""
        self.clear_window()
        
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Title
        title = ttk.Label(main_frame, text="📋 ATTENDANCE SYSTEM", style='Title.TLabel')
        title.pack(pady=20)
        
        # Status label showing loaded students
        status_ok, status_msg = self.student_registry.load_students()
        status_color = "Success.TLabel" if status_ok else "Error.TLabel"
        status = ttk.Label(main_frame, text=f"✓ {status_msg}" if status_ok else f"✗ {status_msg}", style=status_color)
        status.pack(pady=10)
        
        # Button frame
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=40)
        
        btn_new = ttk.Button(
            btn_frame, 
            text="➕ CREATE NEW ATTENDANCE", 
            command=self.show_create_sheet,
            style='Action.TButton',
            width=30
        )
        btn_new.pack(pady=10)
        
        btn_load = ttk.Button(
            btn_frame, 
            text="📂 LOAD EXISTING ATTENDANCE", 
            command=self.show_load_sheet,
            style='Action.TButton',
            width=30
        )
        btn_load.pack(pady=10)
        
        btn_quit = ttk.Button(
            btn_frame, 
            text="❌ EXIT", 
            command=self.root.quit,
            style='Action.TButton',
            width=30
        )
        btn_quit.pack(pady=10)
    
    def show_create_sheet(self):
        """Show form to create new attendance sheet."""
        self.clear_window()
        
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Title
        title = ttk.Label(main_frame, text="📝 NEW ATTENDANCE SHEET", style='Title.TLabel')
        title.pack(pady=20)
        
        # Form frame
        form_frame = ttk.Frame(main_frame)
        form_frame.pack(fill=tk.X, pady=20)
        
        fields = {}
        field_names = ["Staff Name", "Staff Student Number", "Event Name", "Venue", "Start Time (HH:MM)", "End Time (HH:MM)"]
        
        for field_name in field_names:
            row = ttk.Frame(form_frame)
            row.pack(fill=tk.X, pady=8)
            
            label = ttk.Label(row, text=f"{field_name}:", style='Normal.TLabel', width=25)
            label.pack(side=tk.LEFT)
            
            entry = ttk.Entry(row, width=40)
            entry.pack(side=tk.LEFT, padx=10)
            fields[field_name] = entry
        
        # Button frame
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=20)
        
        def create():
            values = {k: v.get().strip() for k, v in fields.items()}
            
            if not all(values.values()):
                messagebox.showerror("Error", "All fields are required!")
                return
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"attendance_{values['Event Name'].replace(' ', '_')}_{timestamp}.xlsx"
            
            self.current_sheet = AttendanceSheet(
                filename,
                values['Staff Name'],
                values['Staff Student Number'],
                values['Event Name'],
                values['Venue'],
                values['Start Time (HH:MM)'],
                values['End Time (HH:MM)']
            )
            
            self.show_scanning_interface()
        
        btn_create = ttk.Button(btn_frame, text="✓ START SCANNING", command=create, style='Action.TButton', width=30)
        btn_create.pack(side=tk.LEFT, padx=10)
        
        btn_back = ttk.Button(btn_frame, text="← BACK", command=self.show_menu, style='Action.TButton', width=30)
        btn_back.pack(side=tk.LEFT, padx=10)
    
    def show_loading_interface(self, filename):
        """Show scanning interface for loaded attendance sheet."""
        try:
            wb = openpyxl.load_workbook(filename)
            ws = wb.active
            
            # Extract event info from header
            staff_name = str(ws['A2'].value or "").replace("Staff Name: ", "")
            staff_number = str(ws['A3'].value or "").replace("Staff Student Number: ", "")
            event_name = str(ws['A4'].value or "").replace("Event: ", "")
            venue = str(ws['A5'].value or "").replace("Venue: ", "")
            start_time = str(ws['A6'].value or "").replace("Start Time: ", "")
            end_time = str(ws['A7'].value or "").replace("End Time: ", "")
            
            self.current_sheet = AttendanceSheet(filename, staff_name, staff_number, event_name, venue, start_time, end_time)
            
            # Load existing records
            for row in ws.iter_rows(min_row=11, values_only=True):
                if row[0] and row[1] and row[2]:  # timestamp, rfid, student_id
                    self.current_sheet.records[row[2]] = (row[0], row[1])
            
            self.show_scanning_interface()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")
            self.show_load_sheet()
    
    def show_load_sheet(self):
        """Show list of existing attendance sheets."""
        self.clear_window()
        
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        title = ttk.Label(main_frame, text="📂 LOAD ATTENDANCE", style='Title.TLabel')
        title.pack(pady=20)
        
        # Find attendance files
        attendance_files = sorted([f for f in os.listdir('.') if f.startswith('attendance_') and f.endswith('.xlsx')], reverse=True)
        
        if not attendance_files:
            msg = ttk.Label(main_frame, text="No attendance files found.", style='Info.TLabel')
            msg.pack(pady=20)
            
            btn_back = ttk.Button(main_frame, text="← BACK", command=self.show_menu, style='Action.TButton', width=30)
            btn_back.pack(pady=20)
            return
        
        # Listbox with files
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=('Helvetica', 11), height=15)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        for file in attendance_files:
            listbox.insert(tk.END, file)
        
        def load_selected():
            selection = listbox.curselection()
            if selection:
                filename = attendance_files[selection[0]]
                self.show_loading_interface(filename)
            else:
                messagebox.showwarning("Warning", "Please select a file!")
        
        # Button frame
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=20)
        
        btn_load = ttk.Button(btn_frame, text="✓ LOAD", command=load_selected, style='Action.TButton', width=30)
        btn_load.pack(side=tk.LEFT, padx=10)
        
        btn_back = ttk.Button(btn_frame, text="← BACK", command=self.show_menu, style='Action.TButton', width=30)
        btn_back.pack(side=tk.LEFT, padx=10)
    
    def show_scanning_interface(self):
        """Show fast, responsive scanning interface."""
        self.clear_window()
        
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Header with event info
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=10, padx=10)
        
        event_info = f"{self.current_sheet.event_name} | {self.current_sheet.venue}"
        title = ttk.Label(header_frame, text=f"🎯 {event_info}", style='Header.TLabel')
        title.pack()
        
        # Scanning input - FOCUSED FOR INSTANT INPUT
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=10, padx=10)
        
        ttk.Label(input_frame, text="Scan RFID or Enter Student ID:", style='Normal.TLabel').pack()
        
        self.scan_input = ttk.Entry(input_frame, font=('Helvetica', 14), width=40)
        self.scan_input.pack(pady=5)
        self.scan_input.focus()
        
        # Status label for feedback
        self.status_label = ttk.Label(main_frame, text="Ready to scan...", style='Info.TLabel')
        self.status_label.pack(pady=5)
        
        # Scanned list with scrollbar
        list_frame = ttk.LabelFrame(main_frame, text=f"Scanned: {len(self.current_sheet.records)}", padding=5)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=10, padx=10)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.scan_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=('Courier', 10), height=15)
        self.scan_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.scan_listbox.yview)
        
        self.populate_scan_list()
        
        # Button frame
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10, padx=10)
        
        btn_save = ttk.Button(btn_frame, text="💾 SAVE & FINISH", command=self.finish_scanning, style='Action.TButton')
        btn_save.pack(side=tk.LEFT, padx=5)
        
        btn_clear = ttk.Button(btn_frame, text="🔄 CLEAR ALL", command=self.clear_all_records, style='Action.TButton')
        btn_clear.pack(side=tk.LEFT, padx=5)
        
        btn_back = ttk.Button(btn_frame, text="← BACK", command=self.show_menu, style='Action.TButton')
        btn_back.pack(side=tk.RIGHT, padx=5)
        
        # Bind Enter key for instant scanning
        self.scan_input.bind('<Return>', self.process_scan)
    
    def process_scan(self, event=None):
        """Process scanned RFID or student ID - instant response."""
        identifier = self.scan_input.get().strip()
        self.scan_input.delete(0, tk.END)
        self.scan_input.focus()
        
        if not identifier:
            return
        
        # Look up student
        student_id = None
        rfid = None
        
        if identifier in self.student_registry.rfid_cache:
            rfid = identifier
            student_id = self.student_registry.get_student_by_rfid(rfid)
        elif identifier in self.student_registry.student_cache:
            student_id = identifier
            rfid = self.student_registry.get_student_by_id(student_id)
        else:
            self.status_label.config(text=f"✗ NOT FOUND: {identifier}", style='Error.TLabel')
            self.scan_input.config(foreground='red')
            self.root.after(1500, lambda: self.scan_input.config(foreground='black'))
            return
        
        # Add record
        timestamp, is_new = self.current_sheet.add_record(rfid, student_id)
        
        if is_new:
            student_name = self.student_registry.get_student_name(student_id)
            name_display = f" ({student_name})" if student_name else ""
            self.status_label.config(text=f"✓ {student_id}{name_display} - {timestamp}", style='Success.TLabel')
            self.populate_scan_list()
            self.scan_input.config(foreground='green')
        else:
            self.status_label.config(text=f"⚠ DUPLICATE: {student_id} already scanned", style='Error.TLabel')
            self.scan_input.config(foreground='orange')
        
        self.root.after(1500, lambda: self.scan_input.config(foreground='black'))
    
    def populate_scan_list(self):
        """Update the scanned list display."""
        self.scan_listbox.delete(0, tk.END)
        count = len(self.current_sheet.records)
        
        for idx, (student_id, (timestamp, rfid)) in enumerate(self.current_sheet.records.items(), 1):
            display = f"{idx:3d}. {student_id:15s} | {timestamp} | {rfid}"
            self.scan_listbox.insert(tk.END, display)
        
        # Scroll to bottom
        self.scan_listbox.see(tk.END)
    
    def clear_all_records(self):
        """Clear all scanned records."""
        if messagebox.askyesno("Confirm", "Clear all scanned records?"):
            self.current_sheet.records.clear()
            self.populate_scan_list()
            self.status_label.config(text="Records cleared", style='Info.TLabel')
    
    def finish_scanning(self):
        """Save attendance and return to menu."""
        if not self.current_sheet.records:
            messagebox.showwarning("Warning", "No records to save!")
            return
        
        success, message = self.current_sheet.save_to_excel()
        
        if success:
            messagebox.showinfo("Success", message)
            self.current_sheet = None
            self.show_menu()
        else:
            messagebox.showerror("Error", message)


def main():
    """Entry point."""
    root = tk.Tk()
    app = AttendanceUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
