#!/usr/bin/env python3
"""
Take attendance (refactored and extended)

Changes from original:
- Adds beep + big tick feedback on camera after successful IN/OUT.
- Adds cooldown for OUT after IN (default 5 minutes).
- Ensures each new day consumes only 2 columns: In-Time and Out-Time.
- Keeps Google Sheets I/O in a single background worker thread.
- Camera thread only enqueues scans and displays UI feedback.
"""

import time
import datetime as dt
import threading
import queue
import json
import openpyxl
import numpy as np
import cv2
from pyzbar.pyzbar import decode
from google.oauth2.service_account import Credentials
import gspread
from gspread.utils import rowcol_to_a1
from gspread.cell import Cell
import os

# ---------------------- CONFIG ----------------------
SERVICE_ACCOUNT_FILE = "attendance.json"
SPREADSHEET_ID = "1NnAM6FtuVrbf3h7dbIwCRHUj4UJ2wDO30_jmq1leGkM"
MASTER_STUDENT_SHEET_TITLE = "Students"
XLSX_FILE = "data.xlsx"
STUDENT_PREFIX = "MERL"

# Camera/frame tuning
FRAME_SKIP = 3

# Worker & debounce tuning
WORKER_FLUSH_INTERVAL = 1.0
SCAN_DEBOUNCE_SECONDS = 1.0  # per-frame duplicate protection (keeps camera responsive)

# OUT cooldown: time (seconds) after IN before OUT may be recorded
OUT_COOLDOWN_SECONDS = 300  # 5 minutes; change to 600 for 10 minutes

# Feedback
BEEP_FILE = "beep.mp3"  # place a short beep file here
FEEDBACK_DISPLAY_MS = 700  # show big tick for 700 ms
FEEDBACK_QUEUE_MAX = 256

# ----------------------------------------------------

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

worker_queue = queue.Queue(maxsize=1000)
shutdown_event = threading.Event()

# Worker -> Camera feedback queue (non-blocking)
FEEDBACK_QUEUE = queue.Queue(maxsize=FEEDBACK_QUEUE_MAX)
# feedback items: ("success", student_name, "IN"|"OUT")
# or ("too_soon", student_name, seconds_remaining)

def load_mappings_from_excel(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path)
    sheet = wb.active
    max_rows = sheet.max_row

    id_to_name = {}
    name_to_row_index = {}
    id_to_section = {}

    header_candidates = {"id", "student", "name"}
    first_cell = str(sheet.cell(1, 1).value).strip().lower() if sheet.cell(1, 1).value else ""
    has_header = any(h in first_cell for h in header_candidates)

    start_row = 2 if has_header else 1
    for i in range(start_row, max_rows + 1):
        id_cell = sheet.cell(i, 1).value
        name_cell = sheet.cell(i, 2).value
        section_cell = sheet.cell(i, 3).value if sheet.max_column >= 3 else None
        if id_cell is None or name_cell is None:
            continue
        id_str = str(id_cell).strip()
        name_str = str(name_cell).strip()
        id_to_name[id_str] = name_str
        id_to_section[id_str] = section_cell
        name_to_row_index[name_str] = i
    return id_to_name, name_to_row_index


class SheetWorker(threading.Thread):
    """
    Handles all Google Sheets I/O.
    Adds day columns as a pair (In-Time, Out-Time).
    Maintains an in-memory map of IN timestamps per row to enforce OUT cooldown reliably.
    """

    def __init__(self, service_file, spreadsheet_id, master_students_title):
        super().__init__(daemon=True)
        self.service_file = service_file
        self.spreadsheet_id = spreadsheet_id
        self.master_students_title = master_students_title
        self.client = None
        self.spreadsheet = None
        self.current_month_title = None
        self.worksheet = None
        self.date_columns = {}  # date_str -> (in_col, out_col)
        self.name_to_row = {}  # name -> row
        self.pending_updates = {}  # (row,col) -> value
        self.lock = threading.Lock()
        # in-memory IN timestamps to check cooldowns without re-reading sheet
        self.row_in_timestamp = {}  # row -> unix timestamp when IN was recorded (or set locally)
        self._connect()

    def _connect(self):
        creds = Credentials.from_service_account_file(self.service_file, scopes=SCOPES)
        self.client = gspread.authorize(creds)
        self.spreadsheet = self.client.open_by_key(self.spreadsheet_id)
        self.ensure_month_sheet(dt.date.today())

    def _write_pending(self):
        with self.lock:
            if not self.pending_updates:
                return
            # construct cell list
            cells = []
            for (r, c), v in list(self.pending_updates.items()):
                cells.append(Cell(r, c, v))
            try:
                if cells:
                    # direct update (single RPC)
                    self.worksheet.update_cells(cells)
            except Exception as e:
                print("Worker: error updating cells:", e)
            finally:
                self.pending_updates.clear()

    def enqueue_update(self, row, col, value):
        with self.lock:
            self.pending_updates[(row, col)] = value

    def ensure_month_sheet(self, date_obj):
        month_title = date_obj.strftime("%B %Y")
        if self.current_month_title == month_title and self.worksheet is not None:
            return
        try:
            ws = self.spreadsheet.worksheet(month_title)
            print(f"Worker: using existing worksheet: {month_title}")
        except Exception:
            print(f"Worker: creating new month worksheet: {month_title}")
            # create with plenty of rows/cols
            ws = self.spreadsheet.add_worksheet(title=month_title, rows="500", cols="50")

            # Write reserved header rows:
            # Row 1: date labels (empty initially)
            # Row 2: sub-headers ("In-Time", "Out-Time" under each date)
            try:
                ws.update_cell(1, 1, "Name")      # header label for column 1
                ws.update_cell(2, 1, "")          # leave sub-header cell blank for Name column
            except Exception:
                pass

            # Try to copy student names from previous month’s worksheet
            prev_ws = None
            try:
                # Get all existing sheets and sort by date
                sheet_titles = [s.title for s in self.spreadsheet.worksheets()]
                month_sheets = []
                for t in sheet_titles:
                    try:
                        parsed = dt.datetime.strptime(t, "%B %Y")
                        month_sheets.append((parsed, t))
                    except Exception:
                        continue
                month_sheets.sort(key=lambda x: x[0])
                prev_ws_title = None
                for d, t in month_sheets:
                    if d < date_obj.replace(day=1):
                        prev_ws_title = t
                if prev_ws_title:
                    prev_ws = self.spreadsheet.worksheet(prev_ws_title)
            except Exception:
                pass

            student_vals = []
            if prev_ws:
                # pull students starting at row 3
                student_vals = prev_ws.col_values(1)[2:]
            if not student_vals:
                # fallback to master sheet
                try:
                    master_ws = self.spreadsheet.worksheet(self.master_students_title)
                except Exception:
                    master_ws = self.spreadsheet.get_worksheet(0)
                student_vals = master_ws.col_values(1)

            # write students starting at row 3
            start_row = 3
            for i, name in enumerate(student_vals, start=start_row):
                if not name:
                    continue
                try:
                    ws.update_cell(i, 1, name)
                except Exception:
                    pass

        self.worksheet = ws
        self.current_month_title = month_title

        # build name->row map
        col1_vals = self.worksheet.col_values(1)
        self.name_to_row = {}
        if col1_vals and str(col1_vals[0]).strip().lower() == "name":
            start_idx = 3
        else:
            start_idx = 1
        for r, n in enumerate(col1_vals, start=1):
            if r < start_idx:
                continue
            if not n:
                continue
            self.name_to_row[str(n).strip()] = r

        # build date_columns mapping
        row1 = self.worksheet.row_values(1)
        self.date_columns = {}
        col = 2
        while col <= len(row1):
            date_label = row1[col - 1] if (col - 1) < len(row1) else None
            if date_label:
                self.date_columns[str(date_label).strip()] = (col, col + 1)
                col += 2
            else:
                break

        # populate row_in_timestamp for today's in-column if present
        today_str = str(dt.date.today())
        if today_str in self.date_columns:
            in_col, out_col = self.date_columns[today_str]
            try:
                col_vals = self.worksheet.col_values(in_col)
                for r, v in enumerate(col_vals, start=1):
                    if v:
                        try:
                            parsed = dt.datetime.strptime(v.strip(), "%I:%M %p")
                            dt_obj = dt.datetime.combine(dt.date.today(), parsed.time())
                            self.row_in_timestamp[r] = dt_obj.timestamp()
                        except Exception:
                            pass
            except Exception:
                pass


    def get_or_create_today_columns(self):
        today_str = str(dt.date.today())
        if today_str in self.date_columns:
            return self.date_columns[today_str]

        # Determine the column to append date columns after the last used date pair.
        # We examine row 1 to find the first empty column slot after existing pairs.
        row1 = self.worksheet.row_values(1)
        # Starting at col 2 (col 1 reserved for Name)
        col = 2
        while True:
            if col - 1 >= len(row1) or not row1[col - 1]:
                # this slot is free: we'll place today's date here (and out-time in next col)
                new_col = col
                break
            col += 2

        try:
            # write the date label in row1, and subheaders in row2
            self.worksheet.update_cell(1, new_col, today_str)
            # put human readable sub-headers
            self.worksheet.update_cell(2, new_col, "In-Time")
            self.worksheet.update_cell(2, new_col + 1, "Out-Time")
        except Exception as e:
            print("Worker: error adding day columns:", e)

        self.date_columns[today_str] = (new_col, new_col + 1)
        return self.date_columns[today_str]

    def process_scan_event(self, student_name, scan_time_iso):
        # ensure sheet exists
        self.ensure_month_sheet(dt.date.today())
        row = self.name_to_row.get(student_name)
        if row is None:
            print(f"Worker: Student '{student_name}' not found; skipping")
            return

        in_col, out_col = self.get_or_create_today_columns()

        # Use in-memory timestamp if available to avoid race
        now_ts = dt.datetime.fromisoformat(scan_time_iso).timestamp()

        # Check current IN and OUT values (sheet) but prefer in-memory if present
        try:
            current_in = self.worksheet.cell(row, in_col).value
        except Exception as e:
            current_in = None

        try:
            current_out = self.worksheet.cell(row, out_col).value
        except Exception:
            current_out = None

        # If no IN recorded -> record IN
        if not current_in and row not in self.row_in_timestamp:
            ts_readable = dt.datetime.fromtimestamp(now_ts).strftime("%I:%M %p")
            print(f"Worker: marking IN for {student_name} row {row} col {in_col} => {ts_readable}")
            self.enqueue_update(row, in_col, ts_readable)
            # update in-memory timestamp immediately for cooldown logic
            self.row_in_timestamp[row] = now_ts
            # enqueue feedback success
            try:
                FEEDBACK_QUEUE.put_nowait(("success", student_name, "IN"))
            except queue.Full:
                pass
            return

        # If IN exists but OUT empty -> attempt to record OUT but respect cooldown
        if (current_in or row in self.row_in_timestamp) and not current_out:
            # determine IN timestamp
            in_ts = None
            if row in self.row_in_timestamp:
                in_ts = self.row_in_timestamp[row]
            else:
                # try to parse current_in (format "%I:%M %p")
                try:
                    parsed = dt.datetime.strptime(current_in.strip(), "%I:%M %p")
                    in_dt = dt.datetime.combine(dt.date.today(), parsed.time())
                    in_ts = in_dt.timestamp()
                    self.row_in_timestamp[row] = in_ts
                except Exception:
                    in_ts = None

            if in_ts is None:
                # cannot determine IN time; to be safe, don't write OUT immediately, record feedback and skip
                try:
                    FEEDBACK_QUEUE.put_nowait(("too_soon", student_name, OUT_COOLDOWN_SECONDS))
                except queue.Full:
                    pass
                return

            elapsed = now_ts - in_ts
            if elapsed < OUT_COOLDOWN_SECONDS:
                # Too soon: compute seconds remaining and notify
                remaining = int(OUT_COOLDOWN_SECONDS - elapsed)
                print(f"Worker: OUT attempted too soon for {student_name}. remaining {remaining}s")
                try:
                    FEEDBACK_QUEUE.put_nowait(("too_soon", student_name, remaining))
                except queue.Full:
                    pass
                return

            # Cooldown satisfied — record OUT
            ts_readable = dt.datetime.fromtimestamp(now_ts).strftime("%I:%M %p")
            print(f"Worker: marking OUT for {student_name} row {row} col {out_col} => {ts_readable}")
            self.enqueue_update(row, out_col, ts_readable)
            # Optionally clear in-memory IN timestamp to allow future INs next day
            try:
                del self.row_in_timestamp[row]
            except KeyError:
                pass
            try:
                FEEDBACK_QUEUE.put_nowait(("success", student_name, "OUT"))
            except queue.Full:
                pass
            return

        # If both already present, ignore
        print(f"Worker: no-op for {student_name}; IN/OUT already present.")
        return

    def run(self):
        last_flush = time.time()
        while not shutdown_event.is_set():
            try:
                task = worker_queue.get(timeout=0.25)
            except queue.Empty:
                # periodic flush
                if time.time() - last_flush >= WORKER_FLUSH_INTERVAL:
                    self._write_pending()
                    last_flush = time.time()
                continue

            if not task:
                worker_queue.task_done()
                continue

            try:
                cmd = task[0]
                if cmd == "process_scan":
                    _, student_name, scan_time_iso = task
                    self.process_scan_event(student_name, scan_time_iso)
                elif cmd == "ensure_month":
                    _, date_obj = task
                    self.ensure_month_sheet(date_obj)
                elif cmd == "flush":
                    self._write_pending()
                elif cmd == "stop":
                    break
                else:
                    print("Worker: unknown task", task)
            except Exception as e:
                print("Worker: exception while processing task:", e)
            finally:
                try:
                    worker_queue.task_done()
                except Exception:
                    pass

            if time.time() - last_flush >= WORKER_FLUSH_INTERVAL:
                self._write_pending()
                last_flush = time.time()

        # final flush
        self._write_pending()
        print("Worker: exiting")


# ---------------------- Camera / Main ----------------------

def play_beep_nonblocking():
    if not os.path.exists(BEEP_FILE):
        return
    try:
        threading.Thread(target=lambda: __playsound(BEEP_FILE), daemon=True).start()
    except Exception:
        pass

def __playsound(path):
    try:
        # import here to avoid mandatory dependency if not using sound
        from playsound import playsound
        playsound(path)
    except Exception:
        pass

def draw_big_tick(frame):
    h, w = frame.shape[:2]
    center_x = int(w * 0.5)
    center_y = int(h * 0.45)

    # Draw a translucent rectangle behind the tick
    overlay = frame.copy()
    alpha = 0.6
    cv2.rectangle(overlay, (center_x - 260, center_y - 200),
                  (center_x + 260, center_y + 200), (0, 0, 0), -1)
    cv2.addWeighted(overlay, alpha, frame, 1 - alpha, 0, frame)

    # Define tick points (like a check mark "✔")
    start_point = (center_x - 100, center_y)       # left start
    mid_point   = (center_x - 40, center_y + 80)   # down slope
    end_point   = (center_x + 120, center_y - 100) # up slope

    # Draw the two lines of the tick
    cv2.line(frame, start_point, mid_point, (0, 255, 0), 18)  # thick green line
    cv2.line(frame, mid_point, end_point, (0, 255, 0), 18)

def draw_text_with_bg(frame, text, org, font, scale, color, thickness, bg_color, padding=5):
    # Get text size
    (w, h), baseline = cv2.getTextSize(text, font, scale, thickness)
    x, y = org

    # Draw filled rectangle as background
    cv2.rectangle(frame,
                  (x - padding, y - h - padding),
                  (x + w + padding, y + baseline + padding),
                  bg_color, -1)

    # Draw text on top
    cv2.putText(frame, text, org, font, scale, color, thickness)


def camera_loop(id_to_name):
    cap = cv2.VideoCapture(0)
    cap.set(3, 640)
    cap.set(4, 480)
    frame_count = 0

    last_scan_time_for_student = {}  # camera-side debounce (seconds)
    feedback_until = 0
    feedback_type = None  # "IN" or "OUT"
    feedback_student = None

    print("starting scan ......")
    try:
        while True:
            ret, frame = cap.read()
            if not ret:
                time.sleep(0.05)
                continue

            frame_count += 1

            # Check feedback queue
            try:
                while True:
                    item = FEEDBACK_QUEUE.get_nowait()
                    if not item:
                        continue
                    kind = item[0]
                    if kind == "success":
                        _, student_name, inout = item
                        feedback_until = time.time() + (FEEDBACK_DISPLAY_MS / 1000.0)
                        feedback_type = inout
                        feedback_student = student_name
                        play_beep_nonblocking()
                    elif kind == "too_soon":
                        _, student_name, remaining = item
                        # short beep to indicate too soon (optional)
                        # play_beep_nonblocking()  # uncomment if you want beep for too-soon
                        feedback_until = time.time() + 0.6
                        feedback_type = "toosoon"
                        feedback_student = student_name
                    FEEDBACK_QUEUE.task_done()
            except queue.Empty:
                pass

            # Process frame decoding every FRAME_SKIP frames
            if frame_count % FRAME_SKIP == 0:
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                qrs = decode(gray)
                for qr in qrs:
                    try:
                        data = qr.data.decode("utf-8").strip()
                    except Exception:
                        continue
                    if not data:
                        continue
                    # Draw bounding box and label immediately
                    pts = np.array([qr.polygon], np.int32).reshape((-1, 1, 2))
                    cv2.polylines(frame, [pts], True, (0, 255, 0), 2)
                    x, y, w, h = qr.rect
                    cv2.putText(frame, data, (x, y - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 2)

                    if not data.startswith(STUDENT_PREFIX):
                        # ignore non-student QR (e.g. LAN)
                        continue

                    student_name = id_to_name.get(data)
                    if not student_name:
                        print("Unknown student id scanned:", data)
                        continue

                    now = time.time()
                    last = last_scan_time_for_student.get(student_name, 0)
                    if now - last < SCAN_DEBOUNCE_SECONDS:
                        # skip quick duplicates at camera level
                        continue
                    last_scan_time_for_student[student_name] = now

                    # enqueue background processing
                    timestamp_iso = dt.datetime.now().isoformat()
                    try:
                        worker_queue.put_nowait(("process_scan", student_name, timestamp_iso))
                    except queue.Full:
                        print("Worker queue full; dropping scan event.")

            # If feedback active, draw big tick overlay
            if time.time() < feedback_until:
                # draw big tick
                if feedback_type == "toosoon":
                    font = cv2.FONT_HERSHEY_SIMPLEX
                    scale = 1.2
                    thickness = 3
                    text_color = (0, 180, 255)
                    bg_color = (0, 0, 0)   # black background

                    y0 = 80
                    dy = int(40 * scale)

                    draw_text_with_bg(frame, "Han Hogaya Bhai Scan!!",
                                    (50, y0), font, scale, text_color, thickness, bg_color)

                    draw_text_with_bg(frame, "Jaa Kaam Pay Lag Ab!!",
                                    (50, y0 + dy), font, scale, text_color, thickness, bg_color)

                else:
                    draw_big_tick(frame)

            cv2.imshow("Result", frame)
            key = cv2.waitKey(1)
            if key == 27:  # ESC
                break

    except KeyboardInterrupt:
        print("Interrupted.")
    finally:
        cap.release()
        cv2.destroyAllWindows()
        shutdown_event.set()
        # tell worker to stop
        try:
            worker_queue.put_nowait(("stop",))
        except Exception:
            pass


def main():
    # load mapping
    id_to_name, name_to_row_index = load_mappings_from_excel(XLSX_FILE)
    if not id_to_name:
        print("No students loaded from Excel. Exiting.")
        return

    # start worker
    worker = SheetWorker(SERVICE_ACCOUNT_FILE, SPREADSHEET_ID, MASTER_STUDENT_SHEET_TITLE)
    worker.start()

    # ensure month exists
    worker_queue.put(("ensure_month", dt.date.today()))

    # start camera loop (blocks)
    camera_loop(id_to_name)

    # wait for worker to finish flush
    worker.join(timeout=5)
    print("Exiting.")


if __name__ == "__main__":
    main()
