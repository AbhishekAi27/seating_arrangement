import pandas as pd
import os
import logging
import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image, Spacer, Flowable
from reportlab.graphics.shapes import Drawing, Rect, Circle, String
from reportlab.lib.units import inch

# ==========================================
# CONFIGURATION
# ==========================================
INPUT_FILE = "input_data.xlsx"
CSV_MAPPING_FILE = "roll-names-mapping.csv"
OUTPUT_DIR = "output"
LOG_FILE = "errors.txt"
PHOTOS_DIR = os.path.join("data", "photos")

# Sheet Names
SHEET_TIMETABLE     = "in_timetable"
SHEET_COURSE_ROLL   = "in_course_roll_mapping"
SHEET_ROLL_NAME     = "in_roll_name_mapping"
SHEET_ROOMS         = "in_room_capacity"

logging.basicConfig(
    handlers=[logging.FileHandler(LOG_FILE, mode='w'), logging.StreamHandler()],
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# --- Helper: Generate Camera Icon Placeholder ---
def get_no_image_drawing(width=55, height=55):
    """
    Creates a vector drawing of a camera inside a box 
    to mimic the 'No Image Available' icon.
    """
    d = Drawing(width, height)
    
    # 1. Outer Box
    d.add(Rect(0, 0, width, height, strokeWidth=1, strokeColor=colors.black, fillColor=None))
    
    # 2. Camera Body parameters
    cx, cy = width / 2, height / 2 + 5
    cam_w, cam_h = 30, 20
    
    # Camera Body Rect
    d.add(Rect(cx - cam_w/2, cy - cam_h/2, cam_w, cam_h, rx=2, ry=2, strokeColor=colors.black, fillColor=None))
    
    # Camera Flash (small box on top)
    flash_w, flash_h = 10, 4
    d.add(Rect(cx - cam_w/4, cy + cam_h/2, flash_w, flash_h, strokeColor=colors.black, fillColor=None))
    
    # Camera Button (small line)
    d.add(Rect(cx + cam_w/4, cy + cam_h/2, 4, 2, strokeColor=colors.black, fillColor=colors.black))

    # Lens (Circle)
    d.add(Circle(cx, cy, 7, strokeColor=colors.black, fillColor=None))
    d.add(Circle(cx, cy, 2, strokeColor=colors.black, fillColor=colors.black)) # Inner dot
    
    # 3. Text "No Image / Available"
    d.add(String(width/2, 12, "No Image", textAnchor="middle", fontSize=6, fontName="Helvetica"))
    d.add(String(width/2, 5, "Available", textAnchor="middle", fontSize=6, fontName="Helvetica"))
    
    return d

class ExamSeatingSystem:
    def __init__(self):
        self.schedule = []
        self.course_enrollments = {}
        self.student_names = {}
        self.rooms = []
        self.allocations = []
        self.room_stats = []

    def clean_text(self, text):
        val = str(text).strip()
        if val.lower() == 'nan' or val == '':
            return ""
        return val

    def load_data(self):
        logging.info(f"Reading input file: {INPUT_FILE}...")
        
        if not os.path.exists(INPUT_FILE):
            logging.error(f"Input file '{INPUT_FILE}' not found.")
            return False

        try:
            # 1. LOAD EXCEL DATA
            xls = pd.ExcelFile(INPUT_FILE)
            
            # --- Timetable ---
            if SHEET_TIMETABLE in xls.sheet_names:
                df_sched = pd.read_excel(xls, sheet_name=SHEET_TIMETABLE)
                df_sched.columns = [str(c).lower().strip() for c in df_sched.columns]
                
                date_col = next((c for c in df_sched.columns if 'date' in c), None)
                day_col = next((c for c in df_sched.columns if 'day' in c), None)
                morn_col = next((c for c in df_sched.columns if 'morning' in c), None)
                eve_col = next((c for c in df_sched.columns if 'evening' in c), None)

                if date_col:
                    for _, row in df_sched.iterrows():
                        raw_date = row[date_col]
                        if pd.isna(raw_date): continue
                        try:
                            if isinstance(raw_date, datetime.datetime):
                                date_obj = raw_date
                            else:
                                date_obj = pd.to_datetime(str(raw_date).split(' ')[0])
                            
                            date_display = date_obj.strftime("%d-%m-%Y") 
                            iso_date = date_obj.strftime("%Y%m%d")
                            day_str = row[day_col] if day_col else date_obj.strftime("%A")
                        except:
                            continue

                        for session_name, col_name in [('Morning', morn_col), ('Evening', eve_col)]:
                            if col_name and pd.notna(row[col_name]):
                                raw_courses = str(row[col_name])
                                if "NO EXAM" in raw_courses.upper() or raw_courses.strip() == "": continue
                                courses = [c.strip() for c in raw_courses.split(';') if c.strip()]
                                if courses:
                                    self.schedule.append({
                                        'Date': date_display,
                                        'IsoDate': iso_date,
                                        'Day': day_str,
                                        'Session': session_name,
                                        'Courses': courses
                                    })

            # --- Enrollments ---
            if SHEET_COURSE_ROLL in xls.sheet_names:
                df_enrol = pd.read_excel(xls, sheet_name=SHEET_COURSE_ROLL)
                df_enrol.columns = [str(c).lower().strip() for c in df_enrol.columns]
                r_col = next((c for c in df_enrol.columns if 'roll' in c), None)
                c_col = next((c for c in df_enrol.columns if 'course' in c), None)

                if r_col and c_col:
                    for _, row in df_enrol.iterrows():
                        c_code = self.clean_text(row[c_col])
                        r_no = self.clean_text(row[r_col])
                        if c_code and r_no:
                            if c_code not in self.course_enrollments: self.course_enrollments[c_code] = []
                            self.course_enrollments[c_code].append(r_no)
            
            for c in self.course_enrollments:
                self.course_enrollments[c].sort()

            # --- Rooms ---
            if SHEET_ROOMS in xls.sheet_names:
                df_rooms = pd.read_excel(xls, sheet_name=SHEET_ROOMS)
                df_rooms.columns = [str(c).lower().strip() for c in df_rooms.columns]
                r_col = next((c for c in df_rooms.columns if 'room' in c), None)
                c_col = next((c for c in df_rooms.columns if 'cap' in c), None)
                
                if r_col and c_col:
                    for _, row in df_rooms.iterrows():
                        self.rooms.append({
                            'Room': self.clean_text(row[r_col]),
                            'Capacity': int(row[c_col]) if pd.notna(row[c_col]) else 0,
                            'filled': 0,
                            'assignments': {}
                        })
                    self.rooms.sort(key=lambda x: -x['Capacity'])

            # 2. LOAD NAMES (FROM CSV OR EXCEL)
            if os.path.exists(CSV_MAPPING_FILE):
                logging.info(f"Loading names from {CSV_MAPPING_FILE}...")
                try:
                    df_csv = pd.read_csv(CSV_MAPPING_FILE)
                    if len(df_csv.columns) >= 2:
                        for _, row in df_csv.iterrows():
                            r_val = self.clean_text(row.iloc[0])
                            n_val = self.clean_text(row.iloc[1])
                            if r_val: self.student_names[r_val] = n_val
                except Exception as e:
                    logging.error(f"Error reading CSV mapping: {e}")

            if SHEET_ROLL_NAME in xls.sheet_names:
                df_names = pd.read_excel(xls, sheet_name=SHEET_ROLL_NAME)
                for _, row in df_names.iterrows():
                    if len(row) < 2: continue
                    r_val = self.clean_text(row.iloc[0])
                    if r_val and r_val not in self.student_names:
                        self.student_names[r_val] = self.clean_text(row.iloc[1])

            return True
        except Exception as e:
            logging.error(f"Error loading data: {e}")
            return False

    def check_clashes(self, courses_in_slot):
        seen_students = {}
        for course in courses_in_slot:
            students = self.course_enrollments.get(course, [])
            for roll in students:
                if roll in seen_students:
                    return True # Clash found
                seen_students[roll] = course
        return False

    def allocate_session(self, slot_info, buffer, mode):
        # Reset rooms
        for r in self.rooms:
            r['filled'] = 0
            r['assignments'] = {}

        courses = slot_info['Courses']
        courses_sorted = sorted(courses, key=lambda c: len(self.course_enrollments.get(c, [])), reverse=True)

        if self.check_clashes(courses):
            logging.warning(f"Clash detected in slot {slot_info['Date']} {slot_info['Session']}")
            return

        for course in courses_sorted:
            students = self.course_enrollments.get(course, [])[:]
            if not students: continue

            for room in self.rooms:
                if not students: break
                
                eff_cap = max(0, room['Capacity'] - buffer)
                subj_limit = eff_cap // 2 if mode == 'sparse' else eff_cap
                space_in_room = eff_cap - room['filled']
                take = min(len(students), space_in_room, subj_limit)

                if take > 0:
                    batch = students[:take]
                    students = students[take:]
                    room['filled'] += take
                    if course not in room['assignments']: room['assignments'][course] = []
                    room['assignments'][course].extend(batch)

                    self.allocations.append({
                        'Date': slot_info['Date'],
                        'IsoDate': slot_info.get('IsoDate', ''),
                        'Day': slot_info.get('Day', ''),
                        'Session': slot_info['Session'],
                        'Course': course,
                        'Room': room['Room'],
                        'Count': take,
                        'Students': batch
                    })
        
        for r in self.rooms:
            self.room_stats.append({
                'Date': slot_info['Date'],
                'Session': slot_info['Session'],
                'Room': r['Room'],
                'Allocated': r['filled'],
                'Free': r['Capacity'] - r['filled']
            })

    def generate_excel_reports(self):
        if not self.allocations: return
        
        data = []
        for x in self.allocations:
            data.append([x['Date'], x['Session'], x['Course'], x['Room'], x['Count'], ";".join(x['Students'])])
        pd.DataFrame(data, columns=["Date", "Session", "Course", "Room", "Count", "Rolls"]).to_excel(os.path.join(OUTPUT_DIR, "overall_seating.xlsx"), index=False)
        pd.DataFrame(self.room_stats).to_excel(os.path.join(OUTPUT_DIR, "room_stats.xlsx"), index=False)

    def generate_attendance_sheets(self):
        logging.info("Generating PDF Attendance Sheets...")
        
        grouped = {}
        for alloc in self.allocations:
            key = (alloc['IsoDate'], alloc['Date'], alloc.get('Day', ''), alloc['Session'], alloc['Room'], alloc['Course'])
            if key not in grouped: grouped[key] = []
            grouped[key].extend(alloc['Students'])

        for (iso_date, disp_date, day, session, room, course), students in grouped.items():
            safe_course = "".join(x for x in course if x.isalnum() or x in " -_")
            filename = os.path.join(OUTPUT_DIR, f"{iso_date}_{session}_{room}_{safe_course}.pdf")
            
            exam_meta = {
                "date": f"{disp_date} ({day})",
                "shift": session,
                "room": room,
                "subject_name": course,
                "count": len(students)
            }
            
            student_objs = []
            for roll in students:
                student_objs.append({
                    "name": self.student_names.get(roll, "Unknown"),
                    "roll": roll,
                    "photo_filename": f"{roll}.jpg"
                })

            self._generate_iitp_pdf(filename, exam_meta, student_objs)

    def _generate_iitp_pdf(self, filename, exam_meta, students):
        # Landscape to fit the 3-column layout nicely
        doc = SimpleDocTemplate(filename, pagesize=landscape(A4), 
                                rightMargin=0.2*inch, leftMargin=0.2*inch, 
                                topMargin=0.3*inch, bottomMargin=0.3*inch)
        
        elements = []
        styles = getSampleStyleSheet()
        
        # --- 1. Header Section ---
        title_style = ParagraphStyle('Title', parent=styles['Heading1'], alignment=1, fontSize=18, spaceAfter=5, fontName='Helvetica-Bold')
        elements.append(Paragraph("IITP Attendance System", title_style))
        
        h_style = ParagraphStyle('H', fontSize=10, fontName='Helvetica-Bold')
        line1 = f"Date: {exam_meta['date']} | Shift: {exam_meta['shift']} | Room No: {exam_meta['room']} | Student count: {exam_meta['count']}"
        line2 = f"Subject: {exam_meta['subject_name']} | Stud Present: {'_'*15} | Stud Absent: {'_'*15}"
        
        header_data = [[Paragraph(line1, h_style)], [Paragraph(line2, h_style)]]
        header_table = Table(header_data, colWidths=[10.5*inch])
        header_table.setStyle(TableStyle([
            ('BOX', (0,0), (-1,-1), 2, colors.black),
            ('INNERGRID', (0,0), (-1,-1), 0.5, colors.black),
            ('LEFTPADDING', (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ('TOPPADDING', (0,0), (-1,-1), 8),
        ]))
        elements.append(header_table)
        elements.append(Spacer(1, 0.1*inch))

        # --- 2. Student Grid Section ---
        
        def create_student_card(student):
            # Check for photo
            img_path = os.path.join(PHOTOS_DIR, student.get('photo_filename', ''))
            
            if os.path.exists(img_path):
                img_obj = Image(img_path)
                img_obj.drawHeight = 0.8 * inch
                img_obj.drawWidth = 0.8 * inch
            else:
                # Use the generated Camera Drawing
                img_obj = get_no_image_drawing(width=55, height=55)

            # Student Details
            s_name = student['name'][:22] 
            details_text = f"""
            <b>{s_name}</b><br/>
            Roll: {student['roll']}<br/>
            Sign: {'_'*16}
            """
            details = Paragraph(details_text, ParagraphStyle('d', fontSize=9, leading=11))
            
            # Card: [ Image | Details ]
            card_data = [[img_obj, details]]
            card_table = Table(card_data, colWidths=[0.9*inch, 2.2*inch], rowHeights=[0.9*inch])
            card_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN', (0,0), (0,0), 'CENTER'),
                ('LEFTPADDING', (0,0), (-1,-1), 2),
                ('RIGHTPADDING', (0,0), (-1,-1), 2),
            ]))
            return card_table

        # Batch into 3 columns
        grid_data = []
        row_data = []
        columns = 3
        
        for student in students:
            row_data.append(create_student_card(student))
            if len(row_data) == columns:
                grid_data.append(row_data)
                row_data = []
        
        if row_data:
            while len(row_data) < columns:
                row_data.append("")
            grid_data.append(row_data)

        if grid_data:
            main_table = Table(grid_data, colWidths=[3.5*inch]*columns)
            main_table.setStyle(TableStyle([
                ('GRID', (0,0), (-1,-1), 1.5, colors.black),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ]))
            elements.append(main_table)

        # --- 3. Footer ---
        elements.append(Spacer(1, 0.3*inch))
        elements.append(Paragraph("Invigilator Name & Signature", ParagraphStyle('inv', alignment=1, fontSize=10)))
        elements.append(Spacer(1, 0.1*inch))
        
        manual_data = [["SI No.", "Name", "Signature"]]
        for _ in range(3):
            manual_data.append(["", "", ""])
            
        manual_table = Table(manual_data, colWidths=[0.8*inch, 4*inch, 4*inch], rowHeights=[0.3*inch]*4)
        manual_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('ALIGN', (0,0), (-1,0), 'LEFT'),
        ]))
        elements.append(manual_table)
        
        try:
            doc.build(elements)
        except Exception as e:
            logging.error(f"Failed to build PDF {filename}: {e}")