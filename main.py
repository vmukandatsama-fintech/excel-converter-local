import time
import os
import sys
import json
import threading
import datetime
import subprocess
import smtplib
import configparser
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# --- THIRD PARTY IMPORTS ---
try:
    import win32com.client
    import pythoncom
    import openpyxl
    from watchdog.observers import Observer
    from watchdog.events import PatternMatchingEventHandler
    from pystray import Icon, Menu, MenuItem
    from PIL import Image, ImageDraw
    from winotify import Notification, audio
except ImportError as e:
    print(f"Missing dependencies: {e}")
    print("Run: pip install winotify")
    sys.exit(1)

def get_base_path():
    if getattr(sys, 'frozen', False): return Path(sys.executable).parent
    return Path(__file__).resolve().parent

def log_message(message):
    try:
        base_path = get_base_path()
        with open(base_path / "activity_log.txt", "a") as f:
            f.write(f"[{datetime.datetime.now()}] {message}\n")
    except: pass
    console_log(message)

CONSOLE_LOCK = threading.Lock()

def console_log(message):
    try:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with CONSOLE_LOCK:
            print(f"[{timestamp}] {message}")
    except Exception:
        pass

# --- DYNAMIC PATH LOGIC ---
def get_root_directory():
    """Get root directory from config.ini or fall back to user profile."""
    try:
        config_path = get_base_path() / "config.ini"
        if config_path.exists():
            config = configparser.ConfigParser()
            config.read(config_path)
            if 'PATHS' in config:
                root_path = config['PATHS'].get('root_path', '')
                if root_path:
                    # Support both absolute paths and network paths
                    return Path(root_path)
    except Exception as e:
        console_log(f"Error reading config path: {e}")
    
    # Fallback to user profile
    return Path(os.environ['USERPROFILE']) / "Premium Leaf Zimbabwe" / "v2SavannahTEST - Collection Vouchers"

ROOT_DIR = get_root_directory()
JSON_DIR = ROOT_DIR / "Json"
TEMPLATE_DIR = ROOT_DIR / "Templates"
WORK_DIR = ROOT_DIR / "Populated Template"
OUTPUT_DIR = ROOT_DIR / "PDF CVs"
TEMPLATE_FILE = TEMPLATE_DIR / "Collection Voucher Template.xlsx"

def initialize_folders():
    """Builds the 4-folder structure if it doesn't exist."""
    for folder in [JSON_DIR, TEMPLATE_DIR, WORK_DIR, OUTPUT_DIR]:
        folder.mkdir(parents=True, exist_ok=True)

def get_versioned_path(base_path):
    """Return a non-conflicting path by appending _01, _02 if needed."""
    if not base_path.exists():
        return base_path
    stem = base_path.stem
    suffix = base_path.suffix
    counter = 1
    while True:
        candidate = base_path.with_name(f"{stem}_{counter:02d}{suffix}")
        if not candidate.exists():
            return candidate
        counter += 1

def archive_processed_json(processing_path):
    """Move processed .processing file to Json/Processed as .json for audit/history."""
    try:
        processing_path = Path(processing_path)
        processed_dir = JSON_DIR / "Processed"
        processed_dir.mkdir(parents=True, exist_ok=True)

        archived_name = f"{processing_path.stem}_processed.json"
        archived_path = get_versioned_path(processed_dir / archived_name)
        os.rename(processing_path, archived_path)
        log_message(f"ARCHIVED JSON: {archived_path.name}")
        return archived_path
    except Exception as e:
        log_message(f"ARCHIVE ERROR: Could not archive {processing_path} - {e}")
        raise

def normalize_excel_value(value):
    """Convert complex JSON values into scalars supported by openpyxl."""
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool, datetime.date, datetime.datetime)):
        return value
    if isinstance(value, dict):
        for key in ("Value", "value", "Name", "name", "Label", "label", "Text", "text", "Id", "id"):
            if key in value and value[key] is not None:
                return normalize_excel_value(value[key])
        return json.dumps(value, ensure_ascii=False)
    if isinstance(value, list):
        return ", ".join(str(normalize_excel_value(v)) for v in value if v is not None)
    return str(value)

def safe_set_cell(ws, cell_ref, raw_value, field_name):
    """Normalize and safely assign a value to an Excel cell."""
    value = normalize_excel_value(raw_value)
    try:
        ws[cell_ref] = value
    except Exception as e:
        log_message(f"EXCEL ASSIGN ERROR [{field_name} @ {cell_ref}]: raw={raw_value!r} normalized={value!r} err={e}")
        raise

def safe_set_cell_rc(ws, row, column, raw_value, field_name):
    """Normalize and safely assign a value using row/column coordinates."""
    value = normalize_excel_value(raw_value)
    try:
        ws.cell(row=row, column=column).value = value
    except Exception as e:
        log_message(f"EXCEL ASSIGN ERROR [{field_name} @ r{row}c{column}]: raw={raw_value!r} normalized={value!r} err={e}")
        raise

class EmailSender:
    """Handles email sending functionality for PDF distribution."""
    
    def __init__(self, config_path=None):
        self.enabled = False
        self.smtp_server = None
        self.smtp_port = 587
        self.sender_email = None
        self.sender_password = None
        self.recipients = []
        self.subject = "New Collection Voucher PDF Generated"
        self.body_template = "Please find the attached PDF."
        
        if config_path is None:
            config_path = get_base_path() / "config.ini"
        
        self.load_config(config_path)
    
    def load_config(self, config_path):
        """Load email configuration from config.ini file."""
        try:
            if not Path(config_path).exists():
                log_message("Config file not found, email disabled")
                return
            
            config = configparser.ConfigParser()
            config.read(config_path)
            
            if 'EMAIL' not in config:
                log_message("No EMAIL section in config, email disabled")
                return
            
            email_config = config['EMAIL']
            
            # Parse boolean value
            self.enabled = email_config.get('enabled', 'false').lower() in ['true', '1', 'yes']
            
            if not self.enabled:
                log_message("Email notifications are disabled in config")
                return
            
            self.smtp_server = email_config.get('smtp_server', '')
            self.smtp_port = int(email_config.get('smtp_port', '587'))
            self.sender_email = email_config.get('sender_email', '')
            self.sender_password = email_config.get('sender_password', '')
            
            # Parse recipients (comma-separated)
            recipients_str = email_config.get('recipients', '')
            self.recipients = [r.strip() for r in recipients_str.split(',') if r.strip()]
            
            self.subject = email_config.get('subject', self.subject)
            self.body_template = email_config.get('body_template', self.body_template)
            
            # Validate configuration
            if not all([self.smtp_server, self.sender_email, self.sender_password, self.recipients]):
                log_message("EMAIL: Incomplete configuration, email disabled")
                self.enabled = False
            else:
                log_message(f"EMAIL: Enabled - Will send to {len(self.recipients)} recipient(s)")
                
        except Exception as e:
            log_message(f"EMAIL: Config error - {e}")
            self.enabled = False
    
    def send_pdf(self, pdf_path, req_no):
        """Send PDF file via email to configured recipients."""
        if not self.enabled:
            return False
        
        try:
            pdf_path = Path(pdf_path)
            if not pdf_path.exists():
                log_message(f"EMAIL: PDF file not found: {pdf_path}")
                return False
            
            # Create message
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = ', '.join(self.recipients)
            msg['Subject'] = self.subject
            
            # Format body with template variables
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            body = self.body_template.format(
                req_no=req_no,
                timestamp=timestamp,
                filename=pdf_path.name
            )
            
            msg.attach(MIMEText(body, 'plain'))
            
            # Attach PDF
            with open(pdf_path, 'rb') as f:
                pdf_attachment = MIMEApplication(f.read(), _subtype='pdf')
                pdf_attachment.add_header('Content-Disposition', 'attachment', 
                                        filename=pdf_path.name)
                msg.attach(pdf_attachment)
            
            # Send email
            log_message(f"EMAIL: Connecting to {self.smtp_server}:{self.smtp_port}")
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.sender_email, self.sender_password)
                server.send_message(msg)
            
            log_message(f"EMAIL: Successfully sent {pdf_path.name} to {len(self.recipients)} recipient(s)")
            return True
            
        except smtplib.SMTPAuthenticationError:
            log_message("EMAIL ERROR: Authentication failed - check credentials")
            return False
        except smtplib.SMTPException as e:
            log_message(f"EMAIL ERROR: SMTP error - {e}")
            return False
        except Exception as e:
            log_message(f"EMAIL ERROR: {e}")
            return False

class SplashScreen:
    def show(self):
        try:
            os.system("cls")
            print("=" * 60)
            print("PLZ CV ENGINE v2 - PRODUCTION READY")
            print(f"Monitoring Path: {ROOT_DIR}")
            
            # Detect mode
            if str(ROOT_DIR).startswith("\\\\") or ":" in str(ROOT_DIR)[:3]:
                if str(ROOT_DIR).startswith("\\\\"):
                    print("Mode: SERVER (Network Monitoring)")
                else:
                    print("Mode: SERVER (Local Shared Path)")
            else:
                print("Mode: USER (OneDrive Sync)")
            
            print("System: Monitoring for Power Automate JSON syncs...")
            print("=" * 60)
        except: pass

class CVAutomator:
    def __init__(self):
        self.status = "Idle"
        self.status_lock = threading.Lock()
        self.icon = None
        self.observer = None
        self.email_sender = EmailSender()

    def set_status(self, message, notify=False):
        with self.status_lock:
            self.status = message
        console_log(f"STATUS: {message}")
        if self.icon:
            self.icon.title = f"PLZ CV Engine - {message}"
            if notify:
                # Use Windows native toast notification instead of pystray
                try:
                    toast = Notification(
                        app_id="PLZ CV Engine",
                        title="PLZ CV Engine",
                        msg=message,
                        duration="short"
                    )
                    toast.show()
                except Exception as e:
                    log_message(f"Notification error: {e}")
                    # Fallback to pystray notification
                    try:
                        self.icon.notify(message, title="PLZ CV Engine")
                    except:
                        pass

    def process_json(self, json_path):
        processing_path = None
        try:
            json_path = Path(json_path)
            self.set_status(f"Detected: {json_path.name}")
            
            # --- CLAIM FILE IMMEDIATELY (Prevent multi-PC conflicts) ---
            # Rename file to .processing to claim ownership
            processing_path = json_path.with_suffix('.processing')
            try:
                # Atomic rename - only one PC will succeed
                os.rename(json_path, processing_path)
                log_message(f"CLAIMED: {json_path.name} (this PC will process it)")
            except (FileNotFoundError, PermissionError, OSError) as e:
                # Another PC already claimed it or file doesn't exist
                log_message(f"SKIPPED: {json_path.name} (already claimed by another PC or missing)")
                return
            
            # --- SAFE READ RETRY LOOP (Fixes 'Line 1 Char 0' error) ---
            data = None
            retries = 10  # Max 10 seconds of waiting
            while retries > 0:
                if processing_path.exists() and os.path.getsize(processing_path) > 0:
                    try:
                        with open(processing_path, 'r') as f:
                            data = json.load(f)
                        break 
                    except (json.JSONDecodeError, PermissionError):
                        pass # File is still being locked or written by OneDrive
                time.sleep(1)
                retries -= 1

            if data is None:
                log_message(f"FAILED: Sync timeout for {processing_path.name}")
                # Clean up the .processing file
                try:
                    os.remove(processing_path)
                except:
                    pass
                return

            payload = data
            if isinstance(data, dict) and isinstance(data.get("body"), str):
                try:
                    payload = json.loads(data.get("body", "{}"))
                    log_message("Detected wrapped payload format; parsed nested 'body' JSON")
                except json.JSONDecodeError as e:
                    log_message(f"Invalid nested body JSON: {e}")
                    payload = data

            req_no = normalize_excel_value(payload.get("FileName", "Unknown"))
            header = payload.get("Header", {})
            lines = payload.get("Lines", [])

            if not TEMPLATE_FILE.exists():
                log_message(f"CRITICAL: Template missing at {TEMPLATE_FILE}")
                return

            # --- POPULATE EXCEL ---
            self.set_status(f"Populating Template for {req_no}...")
            wb = openpyxl.load_workbook(TEMPLATE_FILE)
            ws = wb["CollectionVoucher"] if "CollectionVoucher" in wb.sheetnames else wb.active

            safe_set_cell(ws, "D5", f"COLLECTION VOUCHER : {req_no}", "Header.FileName")
            safe_set_cell(ws, "D9", header.get("Requestor"), "Header.Requestor")
            safe_set_cell(ws, "D10", header.get("Date"), "Header.Date")
            safe_set_cell(ws, "D11", header.get("Approval"), "Header.Approval")
            safe_set_cell(ws, "D12", header.get("Authorization"), "Header.Authorization")
            safe_set_cell(ws, "D13", header.get("Comments"), "Header.Comments")
            
            safe_set_cell(ws, "G9", header.get("Driver"), "Header.Driver")
            safe_set_cell(ws, "G10", header.get("DriverID"), "Header.DriverID")
            safe_set_cell(ws, "G11", header.get("Truck"), "Header.Truck")
            safe_set_cell(ws, "G12", header.get("Trailer"), "Header.Trailer")
            safe_set_cell(ws, "G13", header.get("Farmer"), "Header.Farmer")

            for i, item in enumerate(lines):
                row = 18 + i
                line_value = item.get("Line", item.get("line"))
                uom_value = item.get("UOM", item.get("Uom"))
                item_value = item.get("Item", item.get("Description"))
                requested_value = item.get("Requested", item.get("requested"))
                issue_value = item.get("Issue", item.get("ThisIssue", item.get("Issued")))
                already_issued_value = item.get("AlreadyIssued", item.get("TotalToDate"))
                balance_value = item.get("Balance", item.get("Remaining"))

                safe_set_cell_rc(ws, row, 2, line_value, f"Lines[{i}].Line")
                safe_set_cell_rc(ws, row, 3, uom_value, f"Lines[{i}].UOM")
                safe_set_cell_rc(ws, row, 4, item_value, f"Lines[{i}].Item")
                safe_set_cell_rc(ws, row, 5, requested_value, f"Lines[{i}].Requested")
                safe_set_cell_rc(ws, row, 6, issue_value, f"Lines[{i}].Issue")
                safe_set_cell_rc(ws, row, 7, already_issued_value, f"Lines[{i}].AlreadyIssued")
                safe_set_cell_rc(ws, row, 8, balance_value, f"Lines[{i}].Balance")

            temp_xlsx = WORK_DIR / f"{req_no}.xlsx"
            wb.save(temp_xlsx)
            
            # --- CONVERSION ---
            self.set_status(f"Converting {req_no} to PDF...")
            pythoncom.CoInitialize()
            pdf_path = get_versioned_path(OUTPUT_DIR / f"{req_no}.pdf")
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible, excel.DisplayAlerts = False, False
            wb_api = excel.Workbooks.Open(str(temp_xlsx.absolute()))
            
            try:
                # Targeted Sheet Export
                ws_to_pdf = wb_api.Worksheets("CollectionVoucher")
                ws_to_pdf.ExportAsFixedFormat(0, str(pdf_path.absolute()))
            except:
                wb_api.Worksheets(1).ExportAsFixedFormat(0, str(pdf_path.absolute()))
                
            wb_api.Close(False)
            excel.Quit()
            
            # --- CLEANUP ---
            archive_processed_json(processing_path)
            os.remove(temp_xlsx)
            
            # Open the PDF file directly
            try:
                os.startfile(str(pdf_path.absolute()))
                log_message(f"Opened PDF: {pdf_path.name}")
            except Exception as e:
                log_message(f"Could not open PDF automatically: {e}")
                # Fallback to opening folder if PDF opening fails
                subprocess.Popen(f'explorer "{OUTPUT_DIR}"')
            
            # Send PDF via email if configured
            email_sent = False
            if self.email_sender.enabled:
                self.set_status(f"Sending email for {req_no}...")
                email_sent = self.email_sender.send_pdf(pdf_path, req_no)
            
            # Update status based on email result
            if email_sent:
                self.set_status(f"Done: {pdf_path.name} (Email sent)", notify=True)
            else:
                self.set_status(f"Done: {pdf_path.name}", notify=True)
            
        except Exception as e:
            log_message(f"Processing Error: {e}")
            try:
                if processing_path and Path(processing_path).exists():
                    retry_json = Path(processing_path).with_suffix('.json')
                    os.rename(processing_path, retry_json)
                    log_message(f"RECOVERY: Restored {Path(retry_json).name} for retry")
            except Exception as recovery_error:
                log_message(f"RECOVERY ERROR: Could not restore JSON for retry - {recovery_error}")
            self.set_status("Error (Check Log)", notify=True)
        finally:
            self.set_status("Idle")

    def process_json_from_processing(self, processing_path):
        """Handle recovery of .processing files left from crashes."""
        try:
            processing_path = Path(processing_path)
            log_message(f"RECOVERING: Processing {processing_path.name}")
            
            # Read the file
            with open(processing_path, 'r') as f:
                data = json.load(f)
            
            req_no = data.get("FileName", "Unknown")
            
            # Check if PDF already exists to avoid duplicate work
            if (OUTPUT_DIR / f"{req_no}.pdf").exists():
                log_message(f"SKIPPED RECOVERY: PDF already exists for {req_no}")
                os.remove(processing_path)
                return
            
            # Rename back to .json and let normal processing handle it
            json_path = processing_path.with_suffix('.json')
            os.rename(processing_path, json_path)
            self.process_json(str(json_path))
            
        except Exception as e:
            log_message(f"Recovery failed for {processing_path}: {e}")
            try:
                os.remove(processing_path)
            except:
                pass

    def handle_event(self, event):
        # Get the correct path based on event type
        if hasattr(event, 'dest_path'):
            path = event.dest_path
        else:
            path = event.src_path
        
        # Skip if path is empty or None
        if not path:
            console_log(f"EVENT DETECTED: {event.event_type} - EMPTY PATH (OneDrive sync artifact)")
            return
        
        # Log full details for debugging
        console_log(f"EVENT DETECTED: {event.event_type}")
        console_log(f"  Full path: {path}")
        console_log(f"  Is directory: {event.is_directory}")
        
        # Skip OneDrive temporary files
        if "~RF" in path or ".TMP" in path.upper() or ".tmp" in path:
            console_log(f"  --> IGNORED: OneDrive temporary file")
            return
        
        # Check if it's a JSON file
        if path.lower().endswith(".json") and not event.is_directory:
            console_log(f"  --> PROCESSING: {Path(path).name}")
            # Add small delay to ensure OneDrive sync is complete
            time.sleep(0.5)
            threading.Thread(target=self.process_json, args=(path,), daemon=True).start()
        else:
            console_log(f"  --> IGNORED: {Path(path).name} (not .json or is directory)")

    def run(self):
        initialize_folders()
        
        # Display monitoring info
        console_log("="*60)
        console_log(f"MONITORING FOLDER: {JSON_DIR.absolute()}")
        console_log(f"PDF OUTPUT FOLDER: {OUTPUT_DIR.absolute()}")
        console_log(f"WORK FOLDER: {WORK_DIR.absolute()}")
        console_log(f"Monitoring folder exists: {JSON_DIR.exists()}")
        console_log(f"Monitoring folder is accessible: {os.access(JSON_DIR, os.R_OK)}")
        console_log("="*60)
        
        # 1. Start Watcher
        event_handler = PatternMatchingEventHandler(patterns=["*.json"], ignore_directories=True)
        event_handler.on_created = self.handle_event
        event_handler.on_moved = self.handle_event
        event_handler.on_modified = self.handle_event  # Also watch for modifications (OneDrive)
        
        self.observer = Observer()
        self.observer.schedule(event_handler, str(JSON_DIR), recursive=False)
        self.observer.start()
        console_log("File watcher started successfully!")

        # 2. Process existing backlog
        existing_json = list(JSON_DIR.glob("*.json"))
        if existing_json:
            console_log(f"Found {len(existing_json)} existing JSON file(s) to process")
            for json_file in existing_json:
                console_log(f"  - {json_file.name}")
                threading.Thread(target=self.process_json, args=(str(json_file),), daemon=True).start()
        else:
            console_log("No existing JSON files found in backlog")
        
        # Also process any .processing files left over from crashes
        processing_files = list(JSON_DIR.glob("*.processing"))
        if processing_files:
            console_log(f"Found {len(processing_files)} leftover .processing file(s)")
            for processing_file in processing_files:
                log_message(f"RECOVERING: Found leftover {processing_file.name} from previous crash")
                threading.Thread(target=self.process_json_from_processing, args=(str(processing_file),), daemon=True).start()

        # 3. Tray Icon
        img = Image.new('RGB', (64, 64), (31, 117, 254))
        draw = ImageDraw.Draw(img)
        draw.text((15, 25), "PLZ", fill="white")
        menu = Menu(
            MenuItem("Check Status", lambda i, item: i.notify(self.status, title="PLZ CV Engine")),
            MenuItem("Open PDF Folder", lambda: subprocess.Popen(f'explorer "{OUTPUT_DIR}"')),
            MenuItem("Exit", self.quit_action)
        )
        self.icon = Icon("PLZ", img, "PLZ CV Engine", menu)
        self.icon.run()

    def quit_action(self, icon):
        if self.observer:
            self.observer.stop()
            self.observer.join()
        icon.stop()

if __name__ == "__main__":
    try:
        runtime_mode = "frozen-exe" if getattr(sys, 'frozen', False) else "python-script"
        runtime_entry = sys.executable if getattr(sys, 'frozen', False) else __file__
        console_log(f"RUNTIME: mode={runtime_mode} entry={runtime_entry}")
    except Exception:
        pass
    SplashScreen().show()
    CVAutomator().run()