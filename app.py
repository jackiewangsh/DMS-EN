"""
DMS - Document Management System
Tech Company File Management System
Main entry point
"""
import os
import io
import pymysql
from dbutils.pooled_db import PooledDB
import uuid
import hashlib
import secrets
from datetime import datetime, date
import threading
import pdfplumber
import jieba
import docx
import openpyxl

# LibreOffice conversion global lock (prevent port conflicts from concurrent calls)
_libreoffice_lock = threading.Lock()
import xlrd
import pptx
from functools import wraps

from flask import (
    Flask, render_template, request, redirect, url_for,
    session, jsonify, send_from_directory, send_file, abort,
    flash, make_response, Response
)
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
import shutil
import smtplib
import email.mime.text as mime_text
import email.header
import json as _json

# =====================================================================
# PathConfig
# =====================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
LOG_DIR = os.path.join(BASE_DIR, 'logs')
SYSTEM_DIR = os.path.join(BASE_DIR, 'system')

# Subdirectories
PUBLISHED_DIR = os.path.join(UPLOAD_DIR, 'published')   # Published File
PENDING_DIR = os.path.join(UPLOAD_DIR, 'pending')       # Pending files
OBSOLETE_DIR = os.path.join(UPLOAD_DIR, 'obsolete')      # Obsolete file
WORKFLOW_DIR = os.path.join(UPLOAD_DIR, 'workflow')      # Workflow attachments

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(PUBLISHED_DIR, exist_ok=True)
os.makedirs(PENDING_DIR, exist_ok=True)
os.makedirs(OBSOLETE_DIR, exist_ok=True)
os.makedirs(WORKFLOW_DIR, exist_ok=True)
os.makedirs(LOG_DIR, exist_ok=True)

# Restore progress tracking (for SSE push)
_restore_progress = {}  # user_id -> {phase, current, total, filename}

DB_PATH = os.path.join(DATA_DIR, 'dms.db')
CONFIG_PATH = os.path.join(DATA_DIR, 'config.json')

# =====================================================================
# Flask configuration
# =====================================================================
app = Flask(__name__,
            template_folder=os.path.join(SYSTEM_DIR, 'templates'),
            static_folder=os.path.join(SYSTEM_DIR, 'static'))
# Fixed secret_key: session persists after restart (required for 7x24 production)
_SECRET_KEY_PATH = os.path.join(DATA_DIR, '.secret_key')
if os.path.exists(_SECRET_KEY_PATH):
    with open(_SECRET_KEY_PATH, 'r') as _f:
        app.secret_key = _f.read().strip()
else:
    app.secret_key = secrets.token_hex(32)
    with open(_SECRET_KEY_PATH, 'w') as _f:
        _f.write(app.secret_key)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500MB

# =====================================================================
# Database utilities
# =====================================================================

# MariaDB connection pool config
db_pool = None

def init_db_pool():
    """InitializeMariaDB连接池"""
    global db_pool
    if db_pool is None:
        db_pool = PooledDB(
            creator=pymysql,
            host='localhost',
            user='dms_user',
            password='DmsUser123!',
            database='dms',
            charset='utf8mb4',
            maxconnections=60,
            autocommit=False,
            cursorclass=pymysql.cursors.DictCursor,
            ping=7  # Connection health check: 1=on connect, 2=before cursor, 4=before cursor(retry), 7=all scenarios+auto reconnect
        )
    return db_pool

def get_db():
    """GetData库连接（MariaDB）"""
    if db_pool is None:
        init_db_pool()
    conn = db_pool.connection()
    return conn


def get_db_cursor():
    """GetData库游标（MariaDB）"""
    conn = get_db()
    return conn.cursor()

def dict_from_row(row):
    """WillData库行转For字典（DictCursor模式下行Already经字典，直接Return）"""
    if isinstance(row, dict):
        return row
    return dict(zip(row.keys(), row))

def _int_param(val, default=1):
    """安全解析Pagination参数，非数字Return默认值"""
    try:
        v = int(val)
        return max(1, v) if v > 0 else default
    except (ValueError, TypeError):
        return default

# =====================================================================
# Database initialization
# =====================================================================
def init_db():
    conn = get_db()
    cur = conn.cursor()

    # Users table
    cur.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id VARCHAR(255) PRIMARY KEY,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            password TEXT DEFAULT '',
            email TEXT,
            department TEXT,
            cdsid TEXT,
            status TEXT DEFAULT 'Active',
            role TEXT DEFAULT 'Regular User',
            created_at TEXT,
            user_theme TEXT DEFAULT 'tech-blue',
            must_change_password INTEGER DEFAULT 0
        )
    ''')

    # Migration: add must_change_password column
    try:
        cur.execute("ALTER TABLE users ADD COLUMN must_change_password INTEGER DEFAULT 0")
    except Exception:
        pass  # Column already exists

    # Migration: clear plaintext passwords (empty password field)
    try:
        cur.execute("UPDATE users SET password='' WHERE password != ''")
    except Exception:
        pass

    # Migration: upgrade SHA256 to werkzeug format (mark for password change, auto upgrade on login)
    # Do not upgrade here (cannot reverse plaintext). Auto handle on login.

    # Departments table
    cur.execute('''
        CREATE TABLE IF NOT EXISTS departments (
            id VARCHAR(255) PRIMARY KEY,
            name TEXT UNIQUE NOT NULL,
            sort_order INTEGER DEFAULT 0,
            created_at TEXT
        )
    ''')

    # File types table (global)
    cur.execute('''
        CREATE TABLE IF NOT EXISTS file_types (
            id VARCHAR(255) PRIMARY KEY,
            name TEXT UNIQUE NOT NULL,
            sort_order INTEGER DEFAULT 0,
            created_at TEXT
        )
    ''')

    # Department file types table (each department has its own types)
    cur.execute('''
        CREATE TABLE IF NOT EXISTS dept_file_types (
            id VARCHAR(255) PRIMARY KEY,
            department TEXT NOT NULL,
            name TEXT NOT NULL,
            sort_order INTEGER DEFAULT 0,
            created_at TEXT,
            UNIQUE(department, name)
        )
    ''')

    # Files table
    cur.execute('''
        CREATE TABLE IF NOT EXISTS files (
            id VARCHAR(255) PRIMARY KEY,
            filename TEXT NOT NULL,
            original_name TEXT NOT NULL,
            department TEXT,
            file_type TEXT,
            file_path TEXT NOT NULL,
            file_size INTEGER,
            status TEXT DEFAULT 'Published',
            uploader_id TEXT,
            uploader_name TEXT,
            created_at TEXT,
            updated_at TEXT,
            is_published INTEGER DEFAULT 1,
            obsolete_reason TEXT,
            FOREIGN KEY(uploader_id) REFERENCES users(id)
        )
    ''')

    # Approval workflows table
    cur.execute('''
        CREATE TABLE IF NOT EXISTS workflows (
            id VARCHAR(255) PRIMARY KEY,
            serial_no TEXT UNIQUE NOT NULL,
            title TEXT NOT NULL,
            initiator_id TEXT,
            initiator_name TEXT,
            department TEXT,
            current_node TEXT,
            status TEXT DEFAULT 'In Progress',
            created_at TEXT,
            updated_at TEXT,
            is_closed INTEGER DEFAULT 0,
            closed_at TEXT,
            FOREIGN KEY(initiator_id) REFERENCES users(id)
        )
    ''')

    # Migration: add archive_path column (for 4-level approval archival)
    try:
        cur.execute("ALTER TABLE workflows ADD COLUMN archive_path TEXT")
    except Exception:
        pass  # Column already exists

    # Approval nodes table (5 fixed nodes, fixed order; dept manager supports multiple selection)
    cur.execute('''
        CREATE TABLE IF NOT EXISTS workflow_nodes (
            id VARCHAR(255) PRIMARY KEY,
            workflow_id TEXT NOT NULL,
            node_name TEXT NOT NULL,
            node_order INTEGER,
            is_required INTEGER DEFAULT 1,
            can_multi INTEGER DEFAULT 0,
            approver_ids TEXT,
            approver_names TEXT,
            waiting_approvers TEXT,
            FOREIGN KEY(workflow_id) REFERENCES workflows(id)
        )
    ''')

    # Approval records table (results and notes)
    cur.execute('''
        CREATE TABLE IF NOT EXISTS workflow_records (
            id VARCHAR(255) PRIMARY KEY,
            workflow_id TEXT NOT NULL,
            node_name TEXT,
            approver_id TEXT,
            approver_name TEXT,
            action TEXT,
            comment TEXT,
            created_at TEXT,
            FOREIGN KEY(workflow_id) REFERENCES workflows(id)
        )
    ''')

    # Workflow attachments table (each attachment maps to department and file type)
    cur.execute('''
        CREATE TABLE IF NOT EXISTS workflow_files (
            id VARCHAR(255) PRIMARY KEY,
            workflow_id TEXT NOT NULL,
            file_id TEXT,
            filename TEXT,
            original_name TEXT,
            department TEXT,
            file_type TEXT,
            file_path TEXT,
            file_size INTEGER,
            created_at TEXT,
            is_published INTEGER DEFAULT 0,
            is_deleted INTEGER DEFAULT 0,
            FOREIGN KEY(workflow_id) REFERENCES workflows(id),
            FOREIGN KEY(file_id) REFERENCES files(id)
        )
    ''')

    # Operation logs table
    cur.execute('''
        CREATE TABLE IF NOT EXISTS operation_logs (
            id VARCHAR(255) PRIMARY KEY,
            user_id TEXT,
            username TEXT,
            action TEXT,
            detail TEXT,
            ip_address TEXT,
            created_at TEXT
        )
    ''')

    # System announcements table
    cur.execute('''
        CREATE TABLE IF NOT EXISTS announcements (
            id VARCHAR(255) PRIMARY KEY,
            content TEXT,
            author_id TEXT,
            created_at TEXT
        )
    ''')

    # Main page rich text content table
    cur.execute('''
        CREATE TABLE IF NOT EXISTS dashboard_richtext (
            id INT AUTO_INCREMENT PRIMARY KEY,
            content TEXT,
            updated_by VARCHAR(100),
            updated_at VARCHAR(50)
        )
    ''')


    # Obsolete files records table
    cur.execute('''
        CREATE TABLE IF NOT EXISTS obsolete_files (
            id VARCHAR(255) PRIMARY KEY,
            file_no TEXT,
            filename TEXT,
            original_name TEXT,
            department TEXT,
            file_type TEXT,
            file_path TEXT,
            file_size INTEGER DEFAULT 0,
            uploader_name TEXT,
            obsolete_by TEXT,
            obsolete_at TEXT,
            reason TEXT,
            note TEXT,
            restored_at TEXT,
            restored_by TEXT,
            created_at TEXT,
            updated_at TEXT
        )
    ''')

    # Backup records table
    cur.execute('''
        CREATE TABLE IF NOT EXISTS backup_records (
            id VARCHAR(255) PRIMARY KEY,
            backup_path TEXT,
            backup_type TEXT,
            operator_id TEXT,
            operator_name TEXT,
            created_at TEXT
        )
    ''')

    # Notifications table (approval node notifications)
    cur.execute('''
        CREATE TABLE IF NOT EXISTS notifications (
            id VARCHAR(255) PRIMARY KEY,
            user_id TEXT NOT NULL,
            type TEXT,
            title TEXT,
            content TEXT,
            workflow_id TEXT,
            serial_no TEXT,
            is_read INTEGER DEFAULT 0,
            created_at TEXT,
            FOREIGN KEY(user_id) REFERENCES users(id),
            FOREIGN KEY(workflow_id) REFERENCES workflows(id)
        )
    ''')

    # System config table
    cur.execute('''
        CREATE TABLE IF NOT EXISTS config (
            `key` VARCHAR(255) PRIMARY KEY,
            value TEXT,
            updated_at TEXT
        )
    ''')
    cur.execute('''
        CREATE TABLE IF NOT EXISTS file_shares (
            id INT AUTO_INCREMENT PRIMARY KEY,
            file_id VARCHAR(255) NOT NULL,
            token VARCHAR(255) UNIQUE NOT NULL,
            created_by VARCHAR(100),
            expires_at VARCHAR(50),
            created_at VARCHAR(50) DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    # Initialize default config
    cur.execute("INSERT IGNORE INTO config (`key`, `value`) VALUES ('system_version', 'V1.0.0')")

    # Database field migration (safe add new columns)
    try:
        cur.execute("ALTER TABLE workflow_nodes ADD COLUMN can_multi INTEGER DEFAULT 0")
    except Exception:
        pass  # Column already exists
    try:
        cur.execute("ALTER TABLE workflow_files ADD COLUMN is_deleted INTEGER DEFAULT 0")
    except Exception:
        pass  # Column already exists
    try:
        cur.execute("ALTER TABLE workflow_files ADD COLUMN original_name TEXT")
    except Exception:
        pass  # Column already exists
    try:
        cur.execute("ALTER TABLE workflow_nodes ADD COLUMN waiting_approvers TEXT")
    except Exception:
        pass  # Column already exists

    conn.commit()

    # Insert default admin account
    cur.execute("SELECT id FROM users WHERE username='admin'")
    if not cur.fetchone():
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        cur.execute('''
            INSERT INTO users (id, username, password_hash, password, email, department, cdsid, status, role, created_at)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        ''', (
            str(uuid.uuid4()), 'admin',
            generate_password_hash('admin123'), '',
            'admin@company.com', 'HQ', 'admin', 'Active', 'System Admin', now
        ))
        conn.commit()
        print("[INIT] Default Admin: admin / admin123")

    # Insert default departments and file types
    cur.execute("SELECT COUNT(*) AS cnt FROM departments")
    if cur.fetchone()['cnt'] == 0:
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        default_depts = [
            (str(uuid.uuid4()), '综合部File', 1, now),
            (str(uuid.uuid4()), '公司File', 2, now),
            (str(uuid.uuid4()), 'Finance', 3, now),
            (str(uuid.uuid4()), 'R&D', 4, now),
            (str(uuid.uuid4()), 'Marketing', 5, now),
            (str(uuid.uuid4()), '人力资源部', 6, now),
            (str(uuid.uuid4()), '质量管理部', 7, now),
            (str(uuid.uuid4()), 'IT', 8, now),
            (str(uuid.uuid4()), '生产部', 9, now),
        ]
        cur.executemany('INSERT INTO departments VALUES (%s,%s,%s,%s)', default_depts)

    cur.execute("SELECT COUNT(*) AS cnt FROM file_types")
    if cur.fetchone()['cnt'] == 0:
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        default_types = [
            (str(uuid.uuid4()), 'Manual', 1, now),
            (str(uuid.uuid4()), 'Procedure', 2, now),
            (str(uuid.uuid4()), 'Work Instruction', 3, now),
            (str(uuid.uuid4()), 'Form', 4, now),
            (str(uuid.uuid4()), 'External Documents', 5, now),
            (str(uuid.uuid4()), 'Record', 6, now),
            (str(uuid.uuid4()), 'Others', 7, now),
        ]
        cur.executemany('INSERT INTO file_types VALUES (%s,%s,%s,%s)', default_types)
        conn.commit()

    # Pre-fill default data only when dept_file_types is empty (avoid overwriting user changes)
    cur.execute("SELECT COUNT(*) AS cnt FROM dept_file_types")
    if cur.fetchone()['cnt'] == 0:
        cur.execute("SELECT id, name FROM departments")
        departments = cur.fetchall()
        cur.execute("SELECT id, name FROM file_types ORDER BY sort_order")
        global_types = cur.fetchall()
        for dept in departments:
            for i, ft in enumerate(global_types, 1):
                cur.execute(
                    "INSERT IGNORE INTO dept_file_types (id, department, name, sort_order, created_at) VALUES (%s,%s,%s,%s,%s)",
                    (str(uuid.uuid4()), dept['name'], ft['name'], i, datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
                )
        conn.commit()

    # Auto migration: add obsolete_reason field
    try:
        cur.execute("SHOW COLUMNS FROM files")  # MariaDB syntax
        columns = [col['Field'] for col in cur.fetchall()]
        if 'obsolete_reason' not in columns:
            cur.execute("ALTER TABLE files ADD COLUMN obsolete_reason TEXT")
            conn.commit()
            print("[MIGRATE] Already添加 obsolete_reason field")
    except Exception as e:
        print(f"[MIGRATE] 迁移Check: {e}")


    # file_index table - file content index (MariaDB)
    cur.execute("""CREATE TABLE IF NOT EXISTS file_index (
        id INT AUTO_INCREMENT PRIMARY KEY,
        file_id VARCHAR(255) NOT NULL,
        file_path VARCHAR(500),
        original_name VARCHAR(255),
        content LONGTEXT,
        file_ext VARCHAR(50),
        file_size BIGINT DEFAULT 0,
        indexed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        UNIQUE KEY uk_file_id (file_id),
        FULLTEXT KEY ft_content (content)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
    """)

    # keyword_index table - keyword inverted index (MariaDB)
    cur.execute("""CREATE TABLE IF NOT EXISTS keyword_index (
        id INT AUTO_INCREMENT PRIMARY KEY,
        term VARCHAR(255) NOT NULL,
        file_id VARCHAR(255) NOT NULL,
        frequency INT DEFAULT 1,
        indexed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        UNIQUE KEY uk_term_file (term, file_id),
        INDEX idx_term (term)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
    """)

    # Performance index: optimize file search for million-level records
    performance_indexes = [
        ('idx_files_status_created', 'files', '(status, created_at)'),
        ('idx_files_dept_type', 'files', '(department, file_type)'),
        ('idx_files_dept_type_status', 'files', '(department, file_type, status)'),
        ('idx_files_uploader', 'files', '(uploader_id, created_at)'),
        ('idx_files_published', 'files', '(is_published, status)'),
        ('idx_logs_created', 'operation_logs', '(created_at)'),
        ('idx_notifications_user_read', 'notifications', '(user_id, is_read)'),
        ('idx_workflows_status', 'workflows', '(status, is_closed, updated_at)'),
        ('idx_workflow_nodes_wf_node', 'workflow_nodes', '(workflow_id, node_name)'),
        ('idx_workflow_records_wf', 'workflow_records', '(workflow_id, created_at)'),
    ]
    for idx_name, table, columns in performance_indexes:
        try:
            cur.execute(f'ALTER TABLE {table} ADD INDEX {idx_name} {columns}')
            conn.commit()
        except Exception:
            pass  # Index already exists

    conn.close()
    print("[DB] Data库Initialize完成")

# =====================================================================
# Helper functions
# =====================================================================
def get_sidebar_data(user_role='Regular User', user_id=None, active_dept=None, active_cat=None):
    """Get侧边栏导航Data"""
    conn = get_db()
    cur = conn.cursor()

    # Hide approval menus for read-only users
    if user_role == 'Read-only User':
        hidden_items = ['approvals', 'upload', 'obsolete', 'users', 'dept_type', 'backup', 'logs', 'system_config']
    else:
        hidden_items = []

    # Get departments (sorted)
    cur.execute('SELECT id, name FROM departments ORDER BY sort_order')
    departments = [dict_from_row(r) for r in cur.fetchall()]

    # Count files per department
    cur.execute('''
        SELECT department, COUNT(*) as count
        FROM files
        WHERE status='Published' AND is_published=1
        GROUP BY department
    ''')
    dept_counts = {r['department']: r['count'] for r in cur.fetchall()}

    # Count obsolete files
    cur.execute("SELECT COUNT(*) as count FROM files WHERE status='Obsolete'")
    obsolete_count = cur.fetchone()['count']

    # Count total files
    cur.execute("SELECT COUNT(*) as count FROM files WHERE status='Published' AND is_published=1")
    total_files = cur.fetchone()['count']

    # Count pending workflows
    cur.execute("SELECT COUNT(*) as count FROM workflows WHERE status='In Progress' AND is_closed=0")
    pending_workflows = cur.fetchone()['count']

    # Count files by department and type
    cur.execute('''
        SELECT department, file_type, COUNT(*) as count
        FROM files
        WHERE status='Published' AND is_published=1
        GROUP BY department, file_type
    ''')
    cat_counts = {}
    for r in cur.fetchall():
        key = (r['department'], r['file_type'])
        cat_counts[key] = r['count']

    # Get file types by department (from dept_file_types)
    dept_file_types = {}
    for dept in departments:
        cur.execute(
            'SELECT name FROM dept_file_types WHERE department=%s ORDER BY sort_order',
            (dept['name'],)
        )
        dept_file_types[dept['name']] = [r['name'] for r in cur.fetchall()]

    conn.close()

    return {
        'departments': departments,
        'dept_file_types': dept_file_types,
        'dept_counts': dept_counts,
        'cat_counts': cat_counts,
        'active_dept': active_dept,
        'active_cat': active_cat,
        'obsolete_count': obsolete_count,
        'config': {},
        'total_files': total_files,
        'pending_workflows': pending_workflows,
        'hidden_items': hidden_items,
    }

def get_user_theme(user_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT user_theme, custom_colors FROM users WHERE id=%s', (user_id,))
    row = cur.fetchone()
    conn.close()
    return row['user_theme'] if row else 'tech-blue'

@app.context_processor
def inject_custom_colors():
    """向所Template injectioncustom_colors"""
    cc = ''
    if 'user_id' in session:
        cc = session.get('custom_colors', '') or ''
    return {'custom_colors': cc}

def log_operation(user_id, username, action, detail='', ip=''):
    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
        INSERT INTO operation_logs (id, user_id, username, action, detail, ip_address, created_at)
        VALUES (%s,%s,%s,%s,%s,%s,%s)
    ''', (str(uuid.uuid4()), user_id, username, action, detail, ip,
          datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    conn.commit()
    # Auto cleanup: check every 100 writes (avoid COUNT full table scan)
    _log_counter = getattr(log_operation, '_counter', 0) + 1
    log_operation._counter = _log_counter
    if _log_counter % 100 == 0:
        try:
            # Efficient cleanup using created_at index, keep last 90 days
            cur.execute('DELETE FROM operation_logs WHERE created_at < DATE_SUB(NOW(), INTERVAL 90 DAY)')
            deleted = cur.rowcount
            if deleted > 0:
                conn.commit()
                print(f'[LOG CLEANUP] 清理 {deleted}  90天前Operation Log')
        except Exception as e:
            print(f'[LOG CLEANUP] 清理Failed: {e}')
    conn.close()

def allowed_file_ext(filename):
    ALLOWED = {'.pdf','.doc','.docx','.xls','.xlsx','.ppt','.pptx','.txt',
               '.jpg','.jpeg','.png','.gif','.bmp','.webp','.zip','.7z','.rar'}
    ext = os.path.splitext(filename)[1].lower()
    return ext in ALLOWED

def allowed_compressed_ext(filename):
    """流程Attachment允许Format"""
    ALLOWED = {'.pdf','.doc','.docx','.xls','.xlsx','.ppt','.pptx','.txt','.zip','.7z'}
    ext = os.path.splitext(filename)[1].lower()
    return ext in ALLOWED

def file_ext_to_icon(filename):
    ext = os.path.splitext(filename)[1].lower()
    icons = {
        '.pdf': 'fa-file-pdf', '.doc': 'fa-file-word', '.docx': 'fa-file-word',
        '.xls': 'fa-file-excel', '.xlsx': 'fa-file-excel',
        '.ppt': 'fa-file-powerpoint', '.pptx': 'fa-file-powerpoint',
        '.txt': 'fa-file-alt', '.zip': 'fa-file-archive', '.7z': 'fa-file-archive',
        '.jpg': 'fa-file-image', '.jpeg': 'fa-file-image', '.png': 'fa-file-image',
        '.gif': 'fa-file-image', '.bmp': 'fa-file-image', '.webp': 'fa-file-image',
    }
    return icons.get(ext, 'fa-file-alt')

# =====================================================================
# File index helper functions
# =====================================================================
def extract_text_from_file(file_path, file_ext):
    """FromFile中提取文本Content，Supports PDF/DOCX/XLSX/XLS/PPTX/TXT"""
    text = ""
    file_ext = file_ext.lower()
    
    try:
        if file_ext == '.pdf':
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"
                    # Extract table text
                    tables = page.extract_tables()
                    for table in tables:
                        for row in table:
                            for cell in row:
                                if cell:
                                    text += str(cell) + " "
                        text += "\n"
        elif file_ext == '.docx':
            doc = docx.Document(file_path)
            for para in doc.paragraphs:
                text += para.text + "\n"
            # Extract table text
            for tbl in doc.tables:
                for row in tbl.rows:
                    for cell in row.cells:
                        if cell.text:
                            text += cell.text + " "
                text += "\n"
            # Extract textbox text
            from docx.oxml.ns import qn
            for shape in doc.element.findall('.//' + qn('w:txbxContent')):
                for t in shape.findall('.//' + qn('w:t')):
                    if t.text:
                        text += t.text + " "
                text += "\n"
        elif file_ext == '.xlsx':
            wb = openpyxl.load_workbook(file_path, data_only=True)
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value:
                            text += str(cell.value) + " "
                    text += "\n"
        elif file_ext == '.xls':
            wb = xlrd.open_workbook(file_path)
            for sheet in wb.sheets():
                for row in range(sheet.nrows):
                    for col in range(sheet.ncols):
                        text += str(sheet.cell_value(row, col)) + " "
                    text += "\n"
        elif file_ext == '.pptx':
            prs = pptx.Presentation(file_path)
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text:
                        text += shape.text + "\n"
        elif file_ext == '.doc':
            import olefile
            import struct as _struct
            ole = olefile.OleFileIO(file_path)
            word_stream = ole.openstream('WordDocument').read()
            fc_min = _struct.unpack_from('<I', word_stream, 0x0018)[0]
            fc_mac = _struct.unpack_from('<I', word_stream, 0x001C)[0]
            if 0 < fc_min < fc_mac <= len(word_stream):
                raw = word_stream[fc_min:fc_mac]
                text = raw.decode('utf-16-le', errors='replace')
                text = ''.join(c if c.isprintable() or c in '\n\r\t ' else ' ' for c in text)
            ole.close()
        elif file_ext == '.txt':
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                text = f.read()
    except Exception as e:
        print(f"[WARN] 文本提取Failed {file_path}: {e}")
        text = ""
    
    return text


def tokenize_text(text):
    """对文本进行分词，ReturnKeywordList（去重）"""
    import re
    
    # Remove spaces between Chinese characters
    text = re.sub(r'([\u4e00-\u9fa5])\s+([\u4e00-\u9fa5])', r'\1\2', text)
    # Handle consecutive spaces multiple times
    for _ in range(3):
        text = re.sub(r'([\u4e00-\u9fa5])\s+([\u4e00-\u9fa5])', r'\1\2', text)
    
    # jieba In Progress
    words = jieba.lcut(text)
    
    # Filter: remove spaces, punctuation, numbers, single characters
    keywords = set()
    for word in words:
        word = word.strip()
        # Keep only Chinese (2+ chars) or English words (3+ chars)
        if re.match(r'^[\u4e00-\u9fa5]{2,}$', word):  # Chinese words
            keywords.add(word)
        elif re.match(r'^[a-zA-Z]{3,}$', word):  # English words
            keywords.add(word.lower())
    
    return list(keywords)


def create_file_index(file_id, file_path, original_name, file_ext, file_size):
    """ForFileCreate倒排Index（含死锁重试逻辑）"""
    import traceback, pymysql, datetime, time, random
    
    log_file = open('c:/ai/qc/dms/index_debug.log', 'a', encoding='utf-8')
    log_file.write(f"\n[{datetime.datetime.now()}] create_file_index 开始: {original_name}\n")
    
    try:
        # 1. Extract text
        content = extract_text_from_file(file_path, file_ext)
        if not content:
            log_file.write(f"  [INFO] File无文本Content，跳过Index: {original_name}\n")
            log_file.close()
            return
        
        keywords = tokenize_text(content)
        log_file.write(f"  文本长度: {len(content)}, Keyword数: {len(keywords)}\n")
        
        # 2. Deadlock retry loop (max 5 times)
        max_retries = 5
        success = False
        
        for attempt in range(max_retries):
            log_file.write(f"  尝试 {attempt+1}/{max_retries}: 连接Data库...\n")
            try:
                conn = pymysql.connect(
                    host='localhost',
                    user='dms_user',
                    password='DmsUser123!',
                    database='dms',
                    charset='utf8mb4',
                    cursorclass=pymysql.cursors.DictCursor
                )
            except Exception as e:
                log_file.write(f"  [ERROR] Data库连接Failed: {e}\n")
                if attempt < max_retries - 1:
                    time.sleep(random.uniform(0.2, 0.8))
                    continue
                else:
                    log_file.close()
                    return
            
            try:
                cur = conn.cursor()
                
                # Delete old index
                cur.execute('DELETE FROM keyword_index WHERE file_id=%s', (file_id,))
                cur.execute('DELETE FROM file_index WHERE file_id=%s', (file_id,))
                
                # Insert file_index
                cur.execute('''
                    INSERT INTO file_index (file_id, file_path, original_name, content, file_ext, file_size)
                    VALUES (%s, %s, %s, %s, %s, %s)
                    ON DUPLICATE KEY UPDATE
                        content = VALUES(content),
                        file_ext = VALUES(file_ext),
                        file_size = VALUES(file_size),
                        indexed_at = CURRENT_TIMESTAMP
                ''', (file_id, file_path, original_name, content[:1000000], file_ext, file_size))
                
                # Batch insert keywords (max 200 per batch to avoid large SQL)
                batch_size = 200
                for i in range(0, len(keywords), batch_size):
                    batch = keywords[i:i + batch_size]
                    placeholders = ','.join(['(%s, %s, 1)'] * len(batch))
                    flat_params = []
                    for kw in batch:
                        flat_params.extend([kw, file_id])
                    cur.execute(f'''
                        INSERT INTO keyword_index (term, file_id, frequency)
                        VALUES {placeholders}
                        ON DUPLICATE KEY UPDATE frequency = frequency + VALUES(frequency)
                    ''', flat_params)
                
                conn.commit()
                log_file.write(f"  [SUCCESS] IndexCreateSuccess: {original_name} (Keyword数: {len(keywords)})\n")
                print(f"[INFO] IndexCreateSuccess: {original_name} (Keyword数: {len(keywords)})")
                success = True
                conn.close()
                break
                
            except pymysql.err.OperationalError as e:
                conn.rollback()
                conn.close()
                if e.args[0] == 1213:  # Deadlock
                    log_file.write(f"  [WARN] 死锁 (尝试 {attempt+1}/{max_retries}), {e}, 重试...\n")
                    if attempt < max_retries - 1:
                        time.sleep(random.uniform(0.5, 1.5))
                        continue
                    else:
                        log_file.write(f"  [ERROR] 死锁重试 {max_retries} 次仍Failed: {original_name}\n")
                        print(f"[ERROR] IndexCreateFailed {original_name}: 死锁重试次数用尽")
                else:
                    log_file.write(f"  [ERROR] Data库Error: {e}\n")
                    print(f"[ERROR] IndexCreateFailed {original_name}: {e}")
                    break
                    
            except Exception as e:
                conn.rollback()
                conn.close()
                log_file.write(f"  [ERROR] IndexCreateFailed {original_name}: {e}\n")
                traceback.print_exc(file=log_file)
                print(f"[ERROR] IndexCreateFailed {original_name}: {e}")
                break
        
        if not success:
            log_file.write(f"  [ERROR] IndexCreateFailed（所重试均Failed）: {original_name}\n")
    
    except Exception as e:
        log_file.write(f"  [ERROR] 未知Error: {e}\n")
        traceback.print_exc(file=log_file)
    finally:
        log_file.close()


# =====================================================================
# Login required decorator
# =====================================================================
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            # API requests return JSON, page requests redirect
            if request.is_json or 'application/json' in request.headers.get('Accept', '') or request.path.startswith('/api/'):
                return jsonify({'success': False, 'message': 'Please firstLogin', 'results': [], 'count': 0}), 401
            return redirect(url_for('login'))
        # Force password change: only allow dashboard and password change API
        if session.get('must_change_password'):
            allowed = ['/dashboard', '/api/change-password', '/api/user-theme', '/api/user/avatar', '/logout']
            if request.path not in allowed:
                if request.path.startswith('/api/'):
                    return jsonify({'success': False, 'message': 'Please change your password first'}), 403
                return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated

def role_required(*roles):
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            if 'user_id' not in session:
                # API or AJAX requests return JSON
                if request.is_json or request.path.startswith('/api/') or request.headers.get('X-Requested-With') == 'XMLHttpRequest':
                    resp = jsonify({'success': False, 'message': 'Please firstLogin', 'need_login': True})
                    resp.status_code = 401
                    return resp
                return redirect(url_for('login'))
            if session.get('role') not in roles:
                if request.is_json or request.path.startswith('/api/') or request.headers.get('X-Requested-With') == 'XMLHttpRequest':
                    resp = jsonify({'success': False, 'message': '您Permission执行此Operation'})
                    resp.status_code = 403
                    return resp
                return render_template('error.html',
                    message='您You do not have permission to access this page')
            return f(*args, **kwargs)
        return decorated
    return decorator

# =====================================================================
# Route
# =====================================================================


def _format_size(size):
    if size is None:
        return "0B"
    if size < 1024:
        return str(int(size)) + "B"
    elif size < 1024 * 1024:
        return "{:.1f}KB".format(size / 1024)
    elif size < 1024 * 1024 * 1024:
        return "{:.1f}MB".format(size / 1024 / 1024)
    else:
        return "{:.1f}GB".format(size / 1024 / 1024 / 1024)


@app.route("/api/search/fulltext", methods=["GET"])
@login_required
def api_fulltext_search():
    """全文Search：使用 keyword_index 倒排Index + FULLTEXT Index回退"""
    keyword = request.args.get("keyword", "").strip()
    page = _int_param(request.args.get("page"), 1)
    page_size = _int_param(request.args.get("page_size"), 20)
    page_size = min(page_size, 100)  # Prevent fetching too much data in one request
    offset = (page - 1) * page_size
    if not keyword or len(keyword) < 2:
        return jsonify({"success": False, "message": "Keyword至少2个字符"})
    try:
        conn = get_db()
        cur = conn.cursor()
        
        # Strategy 1: Use keyword_index inverted index (exact match) - SQL pagination, not fetchall
        cur.execute("""
            SELECT COUNT(*) AS cnt
            FROM keyword_index ki
            JOIN file_index fi ON ki.file_id = fi.file_id
            WHERE ki.term = %s
        """, (keyword,))
        kw_count = cur.fetchone()['cnt']
        
        cur.execute("""
            SELECT ki.file_id, ki.frequency, fi.original_name, fi.file_path, fi.file_ext, fi.file_size, f.id as uuid, f.department, f.file_type
            FROM keyword_index ki
            JOIN file_index fi ON ki.file_id = fi.file_id
            LEFT JOIN files f ON ki.file_id = f.filename
            WHERE ki.term = %s
            ORDER BY ki.frequency DESC
            LIMIT %s OFFSET %s
        """, (keyword, page_size * 3, 0))  # Fetch extra for merging with LIKE results
        kw_rows = cur.fetchall()
        found_ids = set(r['file_id'] for r in kw_rows)
        
        # Strategy 2: Use FULLTEXT index (catch missed tokens) - only strategy 1 uncovered
        ft_rows = []
        ft_count = 0
        if len(found_ids) < page_size * 3:
            import re as _re
            if _re.match(r'^[\u4e00-\u9fa5]+$', keyword):
                chars = list(keyword)
                spaced_pattern = '%'.join(chars)
                like_pattern = '%' + spaced_pattern + '%'
            else:
                like_pattern = '%' + keyword + '%'
            
            # Prefer FULLTEXT index (10x faster), fall back to LIKE
            try:
                exclude_clause = ''
                exclude_params = []
                if found_ids:
                    placeholders = ','.join(['%s'] * len(found_ids))
                    exclude_clause = f'AND fi.file_id NOT IN ({placeholders})'
                    exclude_params = list(found_ids)
                
                cur.execute(f"""
                    SELECT fi.file_id, 1 as frequency, fi.original_name, fi.file_path, fi.file_ext, fi.file_size, f.id as uuid, f.department, f.file_type
                    FROM file_index fi
                    LEFT JOIN files f ON fi.file_id = f.filename
                    WHERE MATCH(fi.content) AGAINST(%s IN BOOLEAN MODE)
                    {exclude_clause}
                    ORDER BY fi.file_size ASC
                    LIMIT %s
                """, [keyword] + exclude_params + [page_size * 3])
                ft_rows = cur.fetchall()
            except Exception:
                # Fall back to LIKE when FULLTEXT unavailable (limit rows scanned)
                try:
                    cur.execute(f"""
                        SELECT fi.file_id, 1 as frequency, fi.original_name, fi.file_path, fi.file_ext, fi.file_size, f.id as uuid, f.department, f.file_type
                        FROM file_index fi
                        LEFT JOIN files f ON fi.file_id = f.filename
                        WHERE fi.content LIKE %s
                        {exclude_clause}
                        ORDER BY fi.file_size ASC
                        LIMIT %s
                    """, [like_pattern] + exclude_params + [page_size])
                    ft_rows = cur.fetchall()
                except Exception:
                    ft_rows = []
        
        # Merge and deduplicate
        all_rows = list(kw_rows)
        for r in ft_rows:
            if r['file_id'] not in found_ids:
                all_rows.append(r)
                found_ids.add(r['file_id'])
        
        # Estimate total (strategy 1 exact + strategy 2 at least)
        total = kw_count + len(ft_rows)
        
        # Pagination
        paged_rows = all_rows[offset:offset + page_size]
        
        results = []
        for r in paged_rows:
            results.append({
                "id": r.get("uuid") or r["file_id"],
                "file_id": r["file_id"],
                "file_name": r["original_name"],
                "file_path": r["file_path"],
                "file_ext": r["file_ext"],
                "file_size": r["file_size"] or 0,
                "size_fmt": _format_size(r["file_size"]),
                "department": r["department"],
                "file_type": r["file_type"],
                "frequency": r.get("frequency", 1)
            })
        cur.close()
        conn.close()
        return jsonify({"success": True, "keyword": keyword, "total": total, "page": page, "page_size": page_size, "results": results})
    except Exception as e:
        print("Fulltext search error:", e)
        return jsonify({"success": False, "message": str(e)})


@app.route("/api/index/status", methods=["GET"])
@login_required
def api_index_status():
    try:
        cur = get_db_cursor()
        cur.execute("SELECT COUNT(*) as c, SUM(LENGTH(content)) as s, MAX(indexed_at) as t FROM file_index")
        s = cur.fetchone()
        cur.execute("SELECT COUNT(*) as k FROM keyword_index")
        k = cur.fetchone()
        cur.close()
        ts = s["s"] or 0
        return jsonify({"success": True, "total_files": s["c"] or 0, "content_size": ts, "size_fmt": _format_size(ts), "last_index_time": s["t"] or "", "total_keywords": k["k"] or 0})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


@app.route("/api/index/build", methods=["POST"])
@role_required('System Admin', 'Quality Specialist')
def api_index_build():
    import threading, subprocess, sys, os
    def run():
        try:
            p = os.path.join(os.path.dirname(__file__), "system", "rebuild_index.py")
            subprocess.run([sys.executable, p, "--auto"], capture_output=True, timeout=3600)
        except Exception as e:
            print("Index build error:", e)
    t = threading.Thread(target=run, daemon=True)
    t.start()
    return jsonify({"success": True, "message": "Index重建Already At后台启动"})

@app.route("/api/index/progress", methods=["GET"])
def api_index_progress():
    """GetIndex重建进度"""
    import json, os
    progress_file = os.path.join(os.path.dirname(__file__), 'data', 'rebuild_progress.json')
    if os.path.exists(progress_file):
        try:
            with open(progress_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return jsonify({"success": True, **data})
        except Exception:
            pass
    return jsonify({"success": True, "indexed": 0, "total": 0, "status": "idle", "message": "未At重建"})


# ── Full-text index ────────────────────────────────────────────────────────
def _update_file_index(file_path, original_name, file_id):
    """后台线程安全地Update单FileIndex（提取文本 + 写入 file_index + keyword_index）"""
    import re
    try:
        import jieba
        jieba.setLogLevel(jieba.logging.INFO)
    except Exception:
        pass

    def clean(text):
        if not text:
            return ""
        text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", " ", text)
        text = re.sub(r"[ \t]+", " ", text)
        return "\n".join(ln.strip() for ln in text.splitlines() if ln.strip())

    def get_content(path, ext):
        text = ""
        ext = ext.lower()
        if ext == ".pdf":
            try:
                import pdfplumber
                parts = []
                with pdfplumber.open(path) as pdf:
                    for page in pdf.pages:
                        # Body text
                        t = page.extract_text() or ""
                        if t.strip():
                            parts.append(t.strip())
                        # Table text
                        tables = page.extract_tables()
                        for table in tables:
                            for row in table:
                                for cell in row:
                                    if cell and str(cell).strip():
                                        parts.append(str(cell).strip())
                text = "\n".join(parts)
            except Exception:
                try:
                    from pdfminer.high_level import extract_text
                    text = extract_text(path)
                except Exception:
                    try:
                        from pypdf import PdfReader
                        r = PdfReader(path)
                        for pg in r.pages:
                            text += pg.extract_text() or ""
                    except Exception:
                        pass
        elif ext == ".docx":
            try:
                import docx
                doc = docx.Document(path)
                parts = []
                # Body paragraphs
                for p in doc.paragraphs:
                    if p.text.strip():
                        parts.append(p.text.strip())
                # All text in tables
                for tbl in doc.tables:
                    for row in tbl.rows:
                        for cell in row.cells:
                            if cell.text.strip():
                                parts.append(cell.text.strip())
                text = "\n".join(parts)
            except Exception:
                pass
        elif ext == ".doc":
            # Old Word format, pure Python extraction (no MS Word needed)
            try:
                import olefile
                import struct as _struct
                ole = olefile.OleFileIO(path)
                word_stream = ole.openstream('WordDocument').read()
                fc_min = _struct.unpack_from('<I', word_stream, 0x0018)[0]
                fc_mac = _struct.unpack_from('<I', word_stream, 0x001C)[0]
                if 0 < fc_min < fc_mac <= len(word_stream):
                    raw = word_stream[fc_min:fc_mac]
                    text = raw.decode('utf-16-le', errors='replace')
                    text = ''.join(c if c.isprintable() or c in '\n\r\t ' else ' ' for c in text)
                ole.close()
            except Exception:
                pass
        elif ext in (".xlsx", ".xls"):
            try:
                import openpyxl
                wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                parts = []
                for ws in wb.worksheets:
                    for row in ws.iter_rows(values_only=True):
                        for cell in row:
                            if cell is not None:
                                parts.append(str(cell))
                wb.close()
                text = " ".join(parts)
            except Exception:
                pass
        elif ext in (".pptx", ".ppt"):
            try:
                from pptx import Presentation
                prs = Presentation(path)
                parts = []
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if hasattr(shape, "text") and shape.text.strip():
                            parts.append(shape.text.strip())
                text = "\n".join(parts)
            except Exception:
                pass
        elif ext in (".txt", ".md"):
            for enc in ("utf-8", "gbk", "gb2312"):
                try:
                    with open(path, "r", encoding=enc) as f:
                        text = f.read()
                    break
                except UnicodeDecodeError:
                    continue
        return clean(text)

    def tokenize(text):
        try:
            import jieba
            # Remove spaces between Chinese characters
            for _ in range(3):
                text = re.sub(r'([\u4e00-\u9fa5])\s+([\u4e00-\u9fa5])', r'\1\2', text)
            words = jieba.cut(text, cut_all=False)
            STOP = {"", "", "At", "", "我", "", "和", "就", "Not", "人", "都", "一", "一个",
                    "上", "也", "很", "To", "说", "Need", "去", "你", "Will", "着", "", "看",
                    "好", "自己", "这", "那", "他", "她", "它", "们", "For", "And", "对", "但",
                    "Or", "以", "而", "之", "于", "From", "把", "被", "由", "因", "If", "则",
                    "Can", "Can", "该", "本", "各", "其", "所", "中", "内", "外", "前", "后",
                    "等", "比", "及", "如", "做", "又", "再", "只", "还", "Already", "曾", "Can以",
                    "这个", "那个", "什么", "怎么", "如何", "年", "月", "日", "时", "分", "秒",
                    "Please", "您", "此", "彼", "并", "且", "某", "另", "个", "如下", "File",
                    "管理", "系统", "公司", "Department", "Type", "Name"}
            result = []
            for w in words:
                w = w.strip().lower()
                if len(w) < 2 or w in STOP:
                    continue
                if re.match(r"^[\u4e00-\u9fff]$", w):
                    continue
                result.append(w)
            return result
        except Exception:
            return []

    if not os.path.exists(file_path):
        return
    ext = os.path.splitext(original_name)[1] if original_name else os.path.splitext(file_path)[1]
    size = os.path.getsize(file_path)
    content = get_content(file_path, ext)
    content_snippet = content[:1000000]

    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute("""
            INSERT INTO file_index (file_id, file_path, original_name, content, file_ext, file_size, indexed_at)
            VALUES (%s,%s,%s,%s,%s,%s,NOW())
            ON DUPLICATE KEY UPDATE
                file_path=VALUES(file_path), original_name=VALUES(original_name),
                content=VALUES(content), file_ext=VALUES(file_ext),
                file_size=VALUES(file_size), indexed_at=NOW()
        """, (file_id, file_path, original_name, content_snippet, ext, size))
        if content:
            tokens = tokenize(content)
            tf = {}
            for t in tokens:
                tf[t] = tf.get(t, 0) + 1
            for term, freq in tf.items():
                cur.execute("""
                    INSERT INTO keyword_index (term, file_id, frequency, indexed_at)
                    VALUES (%s,%s,%s,NOW())
                    ON DUPLICATE KEY UPDATE frequency=VALUES(frequency), indexed_at=NOW()
                """, (term, file_id, freq))
        conn.commit()
        print(f"[Index] IndexAlready Updated: {original_name}")
    except Exception as e:
        conn.rollback()
        print(f"[Index] IndexUpdateFailed {original_name}: {e}")
    finally:
        cur.close()
        conn.close()


def _delete_file_index(filename):
    """后台线程安全地Delete单FileIndex（使用 filename 哈希File Name）"""
    try:
        # Use separate connection (avoid get_db() issues in background thread)
        conn = pymysql.connect(
            host='localhost',
            user='dms_user',
            password='DmsUser123!',
            database='dms',
            charset='utf8mb4',
            cursorclass=pymysql.cursors.DictCursor
        )
        cur = conn.cursor()
        cur.execute("DELETE FROM keyword_index WHERE file_id=%s", (filename,))
        cur.execute("DELETE FROM file_index WHERE file_id=%s", (filename,))
        conn.commit()
        cur.close()
        conn.close()
        print(f"[Index] Already DeletedIndex: {filename}")
    except Exception as e:
        print(f"[Index] Delete IndexFailed {filename}: {e}")


@app.route('/')
def index():
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'GET':
        return render_template('login.html')

    # POST: UsernamePasswordLogin
    username = request.form.get('username', '').strip()
    password = request.form.get('password', '')

    if not username or not password:
        return render_template('login.html', error='Please enter Username and Password')

    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
        SELECT id, username, password_hash, role, status, user_theme, custom_colors, must_change_password
        FROM users WHERE LOWER(username)=LOWER(%s)
    ''', (username,))
    user = cur.fetchone()
    conn.close()

    if user and user['status'] == 'Active':
        # Unified password verification logic
        password_valid = False
        need_upgrade_hash = False

        try:
            # 1. Try Werkzeug format (standard)
            password_valid = check_password_hash(user['password_hash'], password)
        except Exception:
            pass
        
        if not password_valid:
            # 2. Try SHA256 format (old imported users)
            sha256_input = hashlib.sha256(password.encode()).hexdigest()
            if user['password_hash'] == sha256_input:
                password_valid = True
                need_upgrade_hash = True  # Need upgrade to werkzeug format

        if password_valid:
            # Auto upgrade old hash format to werkzeug
            if need_upgrade_hash:
                try:
                    cur2 = conn.cursor() if conn else None
                    if cur2 is None:
                        conn2 = get_db()
                        cur2 = conn2.cursor()
                    cur2.execute('UPDATE users SET password_hash=%s WHERE id=%s',
                                 (generate_password_hash(password), user['id']))
                    if conn:
                        conn.commit()
                    else:
                        conn2.commit()
                        conn2.close()
                except Exception:
                    pass  # Upgrade failure does not affect login

            session['user_id'] = user['id']
            session['username'] = user['username']
            session['role'] = user['role']
            session['user_theme'] = user['user_theme']
            session['custom_colors'] = user['custom_colors'] or ''

            # Check if forced password change needed
            if user['must_change_password']:
                session['must_change_password'] = True
                log_operation(user['id'], user['username'], 'Login', 'UserLogin系统（需修改Password）')
                return redirect(url_for('dashboard'))

            log_operation(user['id'], user['username'], 'Login', 'UserLogin系统')
            return redirect(url_for('dashboard'))

    return render_template('login.html', error='UsernameOrPasswordError')

@app.route('/dashboard')
@login_required
def dashboard():
    user_id = session['user_id']
    username = session['username']
    role = session['role']

    conn = get_db()
    cur = conn.cursor()

    # Count user pending approvals - only count workflows requiring user approval
    cur.execute('''
        SELECT COUNT(DISTINCT w.id) AS cnt
        FROM workflows w
        JOIN workflow_nodes n ON n.workflow_id = w.id
        WHERE w.status='In Progress' AND w.is_closed=0 
        AND n.node_name = w.current_node
        AND n.approver_ids LIKE CONCAT('%%', %s, '%%')
    ''', (str(user_id),))
    pending_count = cur.fetchone()['cnt']

    # Latest published files (top 10)
    cur.execute("SELECT id, filename, original_name, department, file_type, file_size, uploader_name, created_at, status FROM files WHERE is_published=1 AND status NOT IN ('Obsolete', 'Delete') ORDER BY created_at DESC LIMIT 10")
    recent_files = [dict_from_row(r) for r in cur.fetchall()]

    # System announcements
    cur.execute('SELECT content, created_at FROM announcements ORDER BY created_at DESC LIMIT 1')
    ann = cur.fetchone()
    announcement = dict_from_row(ann) if ann else None

    # Main page rich text content
    cur.execute('SELECT content FROM dashboard_richtext LIMIT 1')
    rt_row = cur.fetchone()
    dashboard_richtext = rt_row['content'] if rt_row else ''

    conn.close()

    # Format file size
    for f in recent_files:
        size = f.get('file_size', 0) or 0
        if size < 1024:
            f['size_fmt'] = f'{size}B'
        elif size < 1024*1024:
            f['size_fmt'] = f'{size/1024:.1f}KB'
        else:
            f['size_fmt'] = f'{size/1024/1024:.1f}MB'
        f['icon'] = file_ext_to_icon(f['original_name'])

    sidebar = get_sidebar_data(role, user_id)
    sidebar['user_theme'] = session.get('user_theme', 'tech-blue')

    return render_template('dashboard.html',
        pending_count=pending_count,
        recent_files=recent_files,
        announcement=announcement,
        dashboard_richtext=dashboard_richtext,
        sidebar=sidebar,
        page_title='Home'
    )

@app.route('/api/richtext', methods=['GET', 'POST'])
@login_required
def api_richtext():
    """Main page rich text content API"""
    conn = get_db()
    cur = conn.cursor()
    
    if request.method == 'GET':
        # GET: all logged-in users can read
        cur.execute('SELECT content FROM dashboard_richtext LIMIT 1')
        row = cur.fetchone()
        conn.close()
        return jsonify({'success': True, 'content': row['content'] if row else ''})
    
    # POST: save content (admin/quality specialist only)
    if session.get('role') not in ['System Admin', 'Quality Specialist']:
        return jsonify({'success': False, 'message': '无Permission'}), 403
    data = request.get_json() or {}
    content = data.get('content', '')
    
    cur.execute('SELECT id FROM dashboard_richtext LIMIT 1')
    existing = cur.fetchone()
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    user_id = session.get('user_id', '')
    
    if existing:
        cur.execute('UPDATE dashboard_richtext SET content=%s, updated_at=%s, updated_by=%s WHERE id=%s',
                   (content, now, user_id, existing['id']))
    else:
        cur.execute('INSERT INTO dashboard_richtext (content, updated_at, updated_by) VALUES (%s,%s,%s)',
                   (content, now, user_id))
    
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': 'Save Success'})

@app.route('/files')
@login_required
def files():
    """File Listpage面（Pagination，每page最多100 ）"""
    dept_filter = request.args.get('dept', '')
    cat_filter = request.args.get('cat', '')
    search_query = request.args.get('q', '')
    filter_type = request.args.get('filter', 'all')
    page = _int_param(request.args.get('page'), 1)
    per_page = 100
    offset = (page - 1) * per_page
    
    conn = get_db()
    cur = conn.cursor()
    
    # Build query
    where_clauses = ["status != 'Obsolete'"]
    params = []
    
    if dept_filter:
        where_clauses.append("department = %s")
        params.append(dept_filter)
    if cat_filter:
        where_clauses.append("file_type = %s")
        params.append(cat_filter)
    if search_query:
        where_clauses.append("original_name LIKE %s")
        params.append(f'%{search_query}%')
    if filter_type == 'mine':
        where_clauses.append("uploader_id = %s")
        params.append(session['user_id'])
    
    where_sql = ' AND '.join(where_clauses)
    
    # Total
    cur.execute(f"SELECT COUNT(*) AS cnt FROM files WHERE {where_sql}", params)
    total = cur.fetchone()['cnt']
    total_pages = max(1, (total + per_page - 1) // per_page)
    
    cur.execute(f"SELECT * FROM files WHERE {where_sql} ORDER BY created_at DESC LIMIT %s OFFSET %s", params + [per_page, offset])
    files = [dict_from_row(r) for r in cur.fetchall()]
    
    # Format file sizes
    for f in files:
        size = f.get('file_size', 0) or 0
        if size < 1024:
            f['size_fmt'] = f'{size}B'
        elif size < 1024*1024:
            f['size_fmt'] = f'{size/1024:.1f}KB'
        else:
            f['size_fmt'] = f'{size/1024/1024:.1f}MB'
    
    conn.close()
    
    sidebar = get_sidebar_data(session['role'], session['user_id'], dept_filter, cat_filter)
    sidebar['user_theme'] = session.get('user_theme', 'tech-blue')
    
    return render_template('files.html',
        files=files,
        dept_filter=dept_filter,
        cat_filter=cat_filter,
        search_query=search_query,
        filter_type=filter_type,
        sidebar=sidebar,
        page_title='File Library',
        page=page, total=total, total_pages=total_pages, per_page=per_page
    )

@app.route('/logout')
def logout():
    user_id = session.get('user_id', '')
    username = session.get('username', '')
    if user_id:
        log_operation(user_id, username, '退出', 'User退出系统')
    session.clear()
    return redirect(url_for('login'))


@app.route('/api/user/avatar', methods=['POST'])
@login_required
def upload_avatar():
    """Upload UserAvatar"""
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': 'File'})
    f = request.files['file']
    if not f.filename:
        return jsonify({'success': False, 'message': 'Please select Photo File'})
    
    ext = f.filename.rsplit('.', 1)[-1].lower()
    if ext not in ('jpg', 'jpeg', 'png', 'gif', 'webp'):
        return jsonify({'success': False, 'message': '仅Supports jpg/png/gif/webp Format'})
    
    import os, hashlib
    user_id = session['user_id']
    avatar_dir = os.path.join(os.path.dirname(__file__), 'uploads', 'avatars')
    os.makedirs(avatar_dir, exist_ok=True)
    
    # Delete old avatar
    for old_ext in ('jpg', 'jpeg', 'png', 'gif', 'webp'):
        old_path = os.path.join(avatar_dir, f'user_{user_id}.{old_ext}')
        if os.path.exists(old_path):
            os.remove(old_path)
    
    # Save new avatar
    filename = f'user_{user_id}.{ext}'
    filepath = os.path.join(avatar_dir, filename)
    f.save(filepath)
    
    # Update database
    conn = get_db()
    cur = conn.cursor()
    cur.execute('UPDATE users SET avatar=%s WHERE id=%s', (f'/uploads/avatars/{filename}', user_id))
    conn.commit()
    conn.close()
    
    return jsonify({'success': True, 'message': 'Avatar Updated', 'avatar': f'/uploads/avatars/{filename}'})

@app.route('/uploads/avatars/<path:filename>')
@login_required
def serve_avatar(filename):
    """提供Avatar访问"""
    from flask import send_from_directory
    avatar_dir = os.path.join(os.path.dirname(__file__), 'uploads', 'avatars')
    return send_from_directory(avatar_dir, filename)
@app.route('/api/user-theme', methods=['POST'])
@login_required
def set_user_theme():
    """SettingsUser主题（Supports自定义颜色）"""
    data = request.get_json()
    theme = data.get('theme', 'tech-blue')
    custom_colors = data.get('custom_colors')  # {r,g,b}
    
    conn = get_db()
    cur = conn.cursor()
    if custom_colors and isinstance(custom_colors, dict):
        import json
        colors_str = json.dumps({'r': int(custom_colors.get('r',0)), 'g': int(custom_colors.get('g',0)), 'b': int(custom_colors.get('b',0))})
        cur.execute('UPDATE users SET user_theme=%s, custom_colors=%s WHERE id=%s', (theme, colors_str, session['user_id']))
    else:
        cur.execute('UPDATE users SET user_theme=%s, custom_colors=NULL WHERE id=%s', (theme, session['user_id']))
    conn.commit()
    conn.close()
    
    session['user_theme'] = theme
    return jsonify({'success': True, 'message': '主题AlreadySave'})

@app.route('/view/<file_id>')
@login_required
def file_view_inline(file_id):
    """File内联Preview - forAt iframe 中显示，Not触发Download"""
    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT filename, original_name, file_path FROM files WHERE id=%s', (file_id,))
    f = cur.fetchone()
    conn.close()

    if not f:
        abort(404)

    f = dict_from_row(f)

    if not f['file_path'] or not os.path.exists(f['file_path']):
        abort(404)

    # Check MIME type by extension, inline display
    import os as _os_view
    ext = (_os_view.path.splitext(f['original_name'])[1] or '').lower()
    mime_types = {
        '.pdf': 'application/pdf',
        '.png': 'image/png',
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.gif': 'image/gif',
        '.bmp': 'image/bmp',
        '.webp': 'image/webp',
        '.svg': 'image/svg+xml',
        '.txt': 'text/plain',
        '.html': 'text/html',
        '.htm': 'text/html',
        '.xml': 'application/xml',
        '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        '.doc': 'application/msword',
        '.xls': 'application/vnd.ms-excel',
        '.ppt': 'application/vnd.ms-powerpoint',
    }
    mimetype = mime_types.get(ext, 'application/octet-stream')
    # Use send_file streaming, avoid loading large files into memory
    return send_file(f['file_path'], mimetype=mimetype)

@app.route('/download/<file_id>')
@login_required
def file_download(file_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT filename, original_name, file_path FROM files WHERE id=%s', (file_id,))
    f = cur.fetchone()
    conn.close()

    if not f:
        abort(404)

    f = dict_from_row(f)
    file_dir = os.path.dirname(f['file_path'])
    
    # Check if file exists on disk
    if not f['file_path'] or not os.path.exists(f['file_path']):
        # Return friendly error page
        return render_template('error.html', 
                              error_title='File Not存At',
                              error_message=f'File "{f["original_name"]}" 物理File Already丢失，Please联系System Admin。'), 404
    
    # Whether to force download (via attach=1 parameter)
    force_download = request.args.get('attach') == '1'
    from_preview_pdf = request.args.get('pdf') == '1'
    
    if from_preview_pdf:
        # Download LibreOffice-converted PDF from preview page
        import os as _os2
        pdf_name = _os2.path.splitext(f['original_name'])[0] + '.pdf'
        return send_from_directory(file_dir, f['filename'],
                                  as_attachment=True,
                                  download_name=pdf_name)
    
    log_operation(session['user_id'], session['username'], 'Download',
                  f'DownloadFile: {f["original_name"]}')
    
    # Force download: download with original filename
    if force_download:
        return send_from_directory(file_dir, f['filename'],
                                  as_attachment=True,
                                  download_name=f['original_name'])
    # Preview/inline view: keep original filename, browser decides display or download
    return send_from_directory(file_dir, f['filename'],
                              as_attachment=True,
                              download_name=f['original_name'])


# =====================================================================
# Office file preview - convert to PDF using LibreOffice
# =====================================================================
def convert_to_pdf(input_path, output_dir):
    """使用LibreOfficeWillOfficeFile转换ForPDF（带File修改TimeCache）"""
    import subprocess

    # LibreOfficePath
    libreoffice_path = r'C:\Program Files\LibreOffice\program\soffice.exe'
    if not os.path.exists(libreoffice_path):
        libreoffice_path = 'soffice'

    base_name = os.path.splitext(os.path.basename(input_path))[0]
    expected_pdf = os.path.join(output_dir, f'{base_name}.pdf')

    # Cache check: return cached PDF if exists and source unchanged
    if os.path.exists(expected_pdf):
        src_mtime = os.path.getmtime(input_path)
        pdf_mtime = os.path.getmtime(expected_pdf)
        if pdf_mtime >= src_mtime:
            return expected_pdf

    # Queue conversion: serialize to avoid LibreOffice port conflicts
    try:
        with _libreoffice_lock:
            # Check cache again after acquiring lock (another thread may have converted)
            if os.path.exists(expected_pdf):
                src_mtime = os.path.getmtime(input_path)
                pdf_mtime = os.path.getmtime(expected_pdf)
                if pdf_mtime >= src_mtime:
                    return expected_pdf
            result = subprocess.run(
                [libreoffice_path, '--headless', '--convert-to', 'pdf',
                 '--outdir', output_dir, input_path],
                capture_output=True, text=True, timeout=120
            )
        if result.returncode == 0:
            if os.path.exists(expected_pdf):
                return expected_pdf
        return None
    except Exception as e:
        print(f'LibreOffice转换Failed: {e}')
        return None

@app.route('/preview/<file_id>')
@login_required
def file_preview(file_id):
    """File Preview - SupportsPDF直接Preview和OfficeFile转换ForPDFPreview"""
    # Supports URL ?header=0 to hide header toolbar
    # hide_header check moved after dict_from_row (need to handle PDF files)
    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT filename, original_name, file_path FROM files WHERE id=%s', (file_id,))
    f = cur.fetchone()
    conn.close()
    
    if not f:
        abort(404)
    
    f = dict_from_row(f)
    # Hide header: URL parameter header=0 or PDF file
    is_pdf = f['original_name'].lower().endswith('.pdf')
    hide_header = request.args.get('header') == '0' or is_pdf
    # Extract extension from filename (files.file_ext column)
    import os as _os
    file_ext = (_os.path.splitext(f.get('filename') or '')[1] or '').lower()
    
    if not f['file_path'] or not os.path.exists(f['file_path']):
        return render_template('error.html', 
                          error_title='File Not存At',
                          error_message=f'File "{f["original_name"]}" Not存AtOrAlready被Delete')
    
    # Direct preview format
    direct_view_exts = ['.pdf', '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp']
    if file_ext in direct_view_exts:
        if file_ext == '.pdf':
            return render_template('filepreview.html',
                                file_name=f['original_name'],
                                file_id=file_id,
                                file_path=f'/view/{file_id}',
                                file_type='pdf',
                                hide_header=hide_header)
        else:
            return render_template('filepreview.html',
                                file_name=f['original_name'],
                                file_id=file_id,
                                file_path=f'/view/{file_id}',
                                file_type='image',
                                hide_header=hide_header)
    
    # Office file - direct text extraction (millisecond, no external process)
    # Office file - convert to PDF using LibreOffice for preview
    office_exts = ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx']
    if file_ext in office_exts:
        import uuid as _uuid
        temp_dir = os.path.join(os.path.dirname(__file__), 'temp', 'preview')
        os.makedirs(temp_dir, exist_ok=True)
        pdf_path = convert_to_pdf(f['file_path'], temp_dir)
        if pdf_path:
            log_operation(session['user_id'], session['username'], 'Preview',
                          f'PreviewOfficeFile(转PDF): {f["original_name"]}')
            return send_file(pdf_path, mimetype='application/pdf',
                             as_attachment=False)
        # Fall back to text extraction if conversion fails
        text_content = '[PDF转换Failed，PleaseDownload后ViewOr联系AdminCheckLibreOffice]'
        try:
            if file_ext == '.docx':
                from docx import Document
                doc = Document(f['file_path'])
                paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
                text_content = '\n\n'.join(paras)
            elif file_ext == '.xlsx':
                import openpyxl
                wb = openpyxl.load_workbook(f['file_path'], data_only=True, read_only=True)
                parts = []
                for sn in wb.sheetnames:
                    lines = [f'【工作表: {sn}】']
                    ws = wb[sn]
                    for row in ws.iter_rows(values_only=True):
                        row_str = '  |  '.join([str(c) if c is not None else '' for c in row])
                        if row_str.strip():
                            lines.append(row_str)
                    parts.append('\n'.join(lines))
                text_content = '\n\n'.join(parts)
                wb.close()
            elif file_ext == '.xls':
                import xlrd
                wb = xlrd.open_workbook(f['file_path'])
                parts = []
                for i in range(wb.nsheets):
                    ws = wb.sheet_by_index(i)
                    lines = [f'【工作表: {ws.name}】']
                    for r in range(ws.nrows):
                        lines.append('  |  '.join([str(ws.cell_value(r, c)) for c in range(ws.ncols)]))
                    parts.append('\n'.join(lines))
                text_content = '\n\n'.join(parts)
            elif file_ext == '.doc':
                # .doc (Word 97-2003): extract text using olefile pure Python
                try:
                    import olefile
                    import struct as _struct
                    _ole = olefile.OleFileIO(f['file_path'])
                    _word_stream = _ole.openstream('WordDocument').read()
                    _fc_min = _struct.unpack_from('<I', _word_stream, 0x0018)[0]
                    _fc_mac = _struct.unpack_from('<I', _word_stream, 0x001C)[0]
                    if 0 < _fc_min < _fc_mac <= len(_word_stream):
                        _raw = _word_stream[_fc_min:_fc_mac]
                        text_content = _raw.decode('utf-16-le', errors='replace')
                        # Clean non-printable characters, keep newlines and spaces
                        text_content = ''.join(c if c.isprintable() or c in '\n\r\t ' else ' ' for c in text_content)
                        # Split by paragraph (Word uses \r\r or \n\n as separator)
                        import re as _re
                        text_content = _re.sub(r'[\r\n]{2,}', '\n\n', text_content)
                        text_content = _re.sub(r'(?<!\n)\n(?!\n)', '', text_content)
                    else:
                        text_content = '[.docFileContentFor空OrFormatException]'
                    _ole.close()
                except Exception as _de:
                    text_content = f'[.docContent提取Failed: {_de}]'
            elif file_ext == '.pptx':
                from pptx import Presentation
                prs = Presentation(f['file_path'])
                parts = []
                for i, slide in enumerate(prs.slides, 1):
                    lines = [f'【No. {i} page】']
                    for shape in slide.shapes:
                        if hasattr(shape, 'text') and shape.text.strip():
                            lines.append(shape.text.strip())
                    parts.append('\n'.join(lines))
                text_content = '\n\n'.join(parts)
        except ImportError as ie:
            text_content = f'[缺少依赖库: {str(ie)}]'
        except Exception as e:
            text_content = f'[Content提取Failed: {type(e).__name__}: {e}]'
        if not text_content.strip():
            text_content = '[FileFor空，无Can显示Content]'
        return render_template('filepreview.html',
                               file_name=f['original_name'],
                               file_path=f'/download/{file_id}',
                               file_type='text',
                               file_ext=file_ext,
                               content=text_content,
                               hide_header=hide_header)
    
    return render_template('error.html',
                      error_title='Preview not supported',
                      error_message=f'File Type {file_ext} Not Supported in Online Preview')

# =====================================================================
# Open office file in local application (ms-word/ms-excel/ms-powerpoint URI scheme)
# =====================================================================
@app.route('/upload', methods=['GET', 'POST'])
@role_required('System Admin', 'Quality Specialist')
def upload():
    if session.get('role') == 'Read-only User':
        return redirect(url_for('dashboard'))
    if request.method == 'GET':
        sidebar = get_sidebar_data(session['role'], session['user_id'])
        sidebar['user_theme'] = session.get('user_theme', 'tech-blue')
        return render_template('upload.html', sidebar=sidebar, page_title='Upload File')

    files = request.files.getlist('files')
    department = request.form.get('department', '').strip()
    file_type = request.form.get('file_type', '').strip()

    if not department or not file_type:
        return jsonify({'success': False, 'message': 'Please select Department and File Type'})

    dept_dir = os.path.join(PUBLISHED_DIR, department, file_type)
    os.makedirs(dept_dir, exist_ok=True)

    uploaded = []
    for f in files:
        if not f.filename:
            continue
        original_name = f.filename
        ext = os.path.splitext(original_name)[1]
        saved_name = f'{uuid.uuid4().hex}{ext}'
        file_path = os.path.join(dept_dir, saved_name)

        # Check and delete old file in published/ (overwrite upload)
        # Use connection pool, commit first to ensure latest data
        conn_check = get_db()
        try:
            cur_check = conn_check.cursor()
            cur_check.execute('''
                SELECT id, filename, file_path FROM files
                WHERE department=%s AND file_type=%s AND original_name=%s
                  AND is_published=1 AND status != 'Obsolete'
            ''', (department, file_type, original_name))
            existing = cur_check.fetchone()
            if existing:
                # Delete old physical file in published/ directory
                old_path = existing.get('file_path', '')
                if old_path and os.path.isfile(old_path):
                    try:
                        os.remove(old_path)
                    except Exception:
                        pass
                # Delete old file index
                _delete_file_index(existing['filename'])
                # Delete old database record
                cur_check.execute('DELETE FROM files WHERE id=%s', (existing['id'],))
                conn_check.commit()
        finally:
            conn_check.close()

        f.save(file_path)
        file_size = os.path.getsize(file_path)

        file_id = str(uuid.uuid4())
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        conn = get_db()
        try:
            cur = conn.cursor()
            cur.execute('''
                INSERT INTO files (id, filename, original_name, department, file_type,
                    file_path, file_size, status, uploader_id, uploader_name,
                    created_at, updated_at, is_published)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,1)
            ''', (file_id, saved_name, original_name, department, file_type,
                  file_path, file_size, 'Published', session['user_id'],
                  session['username'], now, now))
            conn.commit()
        except Exception as e:
            conn.rollback()
            raise
        finally:
            conn.close()
        uploaded.append(original_name)
        log_operation(session['user_id'], session['username'], 'Upload',
                       f'Upload File: {original_name} -> {department}/{file_type}')
        
        # Create inverted index (async background) - pass saved_name (hash) as file_id to match files.filename
        threading.Thread(target=create_file_index, args=(saved_name, file_path, original_name, ext, file_size), daemon=True).start()

    return jsonify({'success': True, 'message': f'Successfully Uploaded {len(uploaded)} 个File'})

# =====================================================================
# FileObsolete
# =====================================================================
@app.route('/api/files/obsolete', methods=['POST'])
@role_required('System Admin', 'Quality Specialist')
def obsolete_files():
    data = request.get_json()
    file_ids = data.get('file_ids', [])
    reason = data.get('reason', '').strip()
    if not file_ids:
        return jsonify({'success': False, 'message': 'Please select File'})

    conn = get_db()
    try:
        cur = conn.cursor()
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        moved = 0
        processed_ids = []  # RecordSuccessHandleFileID
        
        for fid in file_ids:
            # GetFileInformation
            cur.execute("SELECT filename, department, file_type, original_name FROM files WHERE id=%s", (fid,))
            row = cur.fetchone()
            if not row:
                continue
            filename = row['filename']
            department = row['department']
            file_type = row['file_type']
            original_name = row['original_name']

            # Build source path (locate by department/type)
            src_dir = os.path.join(PUBLISHED_DIR, department, file_type)
            src_path = os.path.join(src_dir, filename)
            # If not in published directory, try pending directory
            if not os.path.exists(src_path):
                src_dir = os.path.join(PENDING_DIR, department, file_type)
                src_path = os.path.join(src_dir, filename)
            # If still not found, search in root directory
            if not os.path.exists(src_path):
                src_path = os.path.join(PUBLISHED_DIR, filename)
                if not os.path.exists(src_path):
                    src_path = os.path.join(PENDING_DIR, filename)

            # Build target path (obsolete directory, organized by department/type)
            dst_dir = os.path.join(OBSOLETE_DIR, department, file_type)
            os.makedirs(dst_dir, exist_ok=True)
            dst_path = os.path.join(dst_dir, filename)

            # Move file
            file_moved = False
            if os.path.exists(src_path):
                try:
                    shutil.move(src_path, dst_path)
                    file_moved = True
                    moved += 1
                except Exception as e:
                    print(f"[WARN] Failed to move file {filename}: {e}")
                    # Update status even if move fails

            # Update database status (save obsolete reason and new file path)
            # Also set is_published to 0 (no longer published)
            if file_moved:
                cur.execute("UPDATE files SET status='Obsolete', obsolete_reason=%s, file_path=%s, is_published=0, updated_at=%s WHERE id=%s", 
                           (reason, dst_path, now, fid))
            else:
                cur.execute("UPDATE files SET status='Obsolete', obsolete_reason=%s, is_published=0, updated_at=%s WHERE id=%s", 
                           (reason, now, fid))
            processed_ids.append((fid, original_name, filename))

        conn.commit()
    except Exception as e:
        conn.rollback()
        raise
    finally:
        conn.close()
    
    # Log after database closed (avoid locking)
    try:
        user_id = session.get('user_id')
        username = session.get('username')
        if user_id and processed_ids:
            for fid, original_name, _ in processed_ids:
                log_operation(user_id, username, 'Obsolete',
                              f'File: {original_name}, Reason: {reason}')
    except Exception as e:
        print(f"[WARN] Log operation failed: {e}")
    
    # Async delete full-text index (background thread, non-blocking)
    try:
        import threading
        for fid, _, fname in processed_ids:
            threading.Thread(target=_delete_file_index, args=(fname,), daemon=True).start()
    except Exception as e:
        print(f"[WARN] Index deletion failed: {e}")
    
    # Async clear preview cache PDF (hashname.pdf)
    try:
        import threading
        preview_dir = os.path.join(os.path.dirname(__file__), 'temp', 'preview')
        if os.path.isdir(preview_dir):
            for _, _, fname in processed_ids:
                base_name = os.path.splitext(fname)[0]  # Get hash name by removing extension
                pdf_path = os.path.join(preview_dir, base_name + '.pdf')
                if os.path.exists(pdf_path):
                    threading.Thread(target=lambda p=pdf_path: os.remove(p), daemon=True).start()
    except Exception as e:
        print(f"[WARN] Preview cache cleanup failed: {e}")
    
    return jsonify({'success': True, 'message': f'Already Obsolete {len(file_ids)} 个File，Success移动 {moved} 个FileToObsoleteDirectory'})

@app.route('/api/files/restore', methods=['POST'])
@role_required('System Admin', 'Quality Specialist')
def restore_files():
    data = request.get_json()
    file_ids = data.get('file_ids', [])
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    restored_count = 0
    restored_files = []  # [(file_id, new_path, original_name)] forIndexUpdate
    conn = get_db()
    try:
        cur = conn.cursor()
        for fid in file_ids:
            # SearchWhenFileInformation
            cur.execute("SELECT original_name, department, file_type, file_path FROM files WHERE id=%s", (fid,))
            row = cur.fetchone()
            if row:
                f = dict_from_row(row)
                orig_name = f['original_name']
                dept = f['department']
                ftype = f['file_type']
                old_path = f['file_path']
                
                # CheckFileWhetherAt obsolete Directory，IfAtRestoreTo published
                if old_path and old_path.startswith(OBSOLETE_DIR):
                    ext = os.path.splitext(orig_name)[1]
                    saved_name = f'{uuid.uuid4().hex}{ext}'
                    dept_dir = os.path.join(PUBLISHED_DIR, dept, ftype)
                    os.makedirs(dept_dir, exist_ok=True)
                    new_path = os.path.join(dept_dir, saved_name)
                    
                    # IfFileAtTo published
                    if os.path.exists(old_path):
                        shutil.copy2(old_path, new_path)
                        cur.execute("UPDATE files SET status='Published', is_published=1, file_path=%s, updated_at=%s WHERE id=%s",
                                   (new_path, now, fid))
                        restored_files.append((fid, new_path, orig_name))
                    else:
                        # File NotAt，UpdateStatus
                        cur.execute("UPDATE files SET status='Published', is_published=1, updated_at=%s WHERE id=%s", (now, fid))
                else:
                    # Already At，UpdateStatus
                    cur.execute("UPDATE files SET status='Published', is_published=1, updated_at=%s WHERE id=%s", (now, fid))
                    restored_files.append((fid, old_path, orig_name))
                
                restored_count += 1
        conn.commit()
    except Exception as e:
        conn.rollback()
        raise
    finally:
        conn.close()

    # Full-text indexUpdate（）
    import threading
    for fid, fpath, fname in restored_files:
        t = threading.Thread(target=_update_file_index, args=(fpath, fname, fid), daemon=True)
        t.start()

    return jsonify({'success': True, 'message': f'Already Restored {restored_count} 个File'})

@app.route('/api/obsolete/upload', methods=['POST'])
@role_required('System Admin', 'Quality Specialist')
def upload_obsolete_files():
    """手工UploadAlready Obsolete fileToObsoleteDirectory"""
    try:
        files = request.files.getlist('files')
        department = request.form.get('department', '').strip()
        file_type = request.form.get('file_type', '').strip()
        reason = request.form.get('reason', '手工UploadObsolete file').strip()
        
        if not files or all(f.filename == '' for f in files):
            return jsonify({'success': False, 'message': 'Please select File'})
        if not department:
            return jsonify({'success': False, 'message': 'Please select Department'})
        if not file_type:
            return jsonify({'success': False, 'message': 'Please select File Type'})
        
        conn = get_db()
        cur = conn.cursor()
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Directory
        dst_dir = os.path.abspath(os.path.join(OBSOLETE_DIR, department, file_type))
        os.makedirs(dst_dir, exist_ok=True)
        
        uploaded = 0
        errors = []
        for f in files:
            if not f or not f.filename:
                continue
            # File Name（）
            original_name = f.filename  # Not secure_filename，
            ext = os.path.splitext(original_name)[1]
            file_id = str(uuid.uuid4())
            filename = f"{file_id}{ext}"
            dst_path = os.path.join(dst_dir, filename)
            
            # SaveFile
            try:
                f.save(dst_path)
                file_size = os.path.getsize(dst_path)
            except Exception as e:
                errors.append(f"{original_name}: {str(e)}")
                continue
            
            # Data
            try:
                cur.execute("""
                    INSERT INTO files (id, filename, original_name, department, file_type,
                                       file_path, file_size, uploader_id, uploader_name, status,
                                       is_published, created_at, updated_at, obsolete_reason)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                """, (file_id, filename, original_name, department, file_type,
                      dst_path, file_size, session['user_id'], session['username'], 'Obsolete',
                      0, now, now, reason))
            except Exception as e:
                errors.append(f"{original_name}: Data库Error - {str(e)}")
                continue
            
            uploaded += 1
        
        conn.commit()
        conn.close()
        
        msg = f'Successfully Uploaded {uploaded} 个Obsolete file'
        if errors:
            msg += f'，Failed {len(errors)} 个'
        return jsonify({'success': True, 'message': msg, 'uploaded': uploaded, 'errors': errors})
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'message': f'Server Error: {str(e)}'})

@app.route('/obsolete')
@role_required('System Admin', 'Quality Specialist')
def obsolete_files_page():
    """Obsolete filepage面 - 只显示Data库Record（Pagination，每page最多100 ）"""
    search = request.args.get('q', '').strip()
    page = _int_param(request.args.get('page'), 1)
    per_page = 100
    offset = (page - 1) * per_page
    conn = get_db()
    cur = conn.cursor()
    
    # Total
    if search:
        cur.execute("""
            SELECT COUNT(*) AS cnt FROM files WHERE status='Obsolete'
              AND (original_name LIKE %s OR department LIKE %s OR file_type LIKE %s OR filename LIKE %s)
        """, (f"%{search}%", f"%{search}%", f"%{search}%", f"%{search}%"))
    else:
        cur.execute("SELECT COUNT(*) AS cnt FROM files WHERE status='Obsolete'")
    total = cur.fetchone()['cnt']
    total_pages = max(1, (total + per_page - 1) // per_page)
    
    if search:
        cur.execute("""
            SELECT id, filename, original_name, department, file_type,
                   file_size, uploader_name, uploader_id, created_at,
                   updated_at, status, is_published, obsolete_reason, file_path
            FROM files WHERE status='Obsolete'
              AND (original_name LIKE %s OR department LIKE %s OR file_type LIKE %s OR filename LIKE %s)
            ORDER BY updated_at DESC
            LIMIT %s OFFSET %s
        """, (f"%{search}%", f"%{search}%", f"%{search}%", f"%{search}%", per_page, offset))
    else:
        cur.execute("""
            SELECT id, filename, original_name, department, file_type,
                   file_size, uploader_name, uploader_id, created_at,
                   updated_at, status, is_published, obsolete_reason, file_path
            FROM files WHERE status='Obsolete'
            ORDER BY updated_at DESC
            LIMIT %s OFFSET %s
        """, (per_page, offset))
    files_list = [dict_from_row(r) for r in cur.fetchall()]
    conn.close()
    
    # FormatFile Size
    for f in files_list:
        sz = f.get('file_size', 0) or 0
        if sz < 1024: f['size_fmt'] = f"{sz}B"
        elif sz < 1024*1024: f['size_fmt'] = f"{sz/1024:.1f}KB"
        else: f['size_fmt'] = f"{sz/1024/1024:.1f}MB"
        f['icon'] = file_ext_to_icon(f.get('original_name', ''))
    
    sidebar = get_sidebar_data(session['role'], session['user_id'],
                               active_dept='__obsolete__', active_cat=None)
    sidebar['user_theme'] = session.get('user_theme', 'tech-blue')
    
    return render_template('obsolete.html', files_list=files_list,
                            search_query=search, sidebar=sidebar, page_title='Obsolete file',
                            page=page, total=total, total_pages=total_pages, per_page=per_page)

@app.route('/api/files/obsolete-list')
@role_required('System Admin', 'Quality Specialist')
def api_obsolete_list():
    search = request.args.get('q', '').strip()
    conn = get_db()
    cur = conn.cursor()
    if search:
        cur.execute("""
            SELECT id, filename, original_name, department, file_type,
                   file_size, uploader_name, uploader_id, created_at, updated_at
            FROM files WHERE status='Obsolete'
              AND (original_name LIKE %s OR department LIKE %s OR file_type LIKE %s)
            ORDER BY updated_at DESC LIMIT 20
        """, (f"%{search}%", f"%{search}%", f"%{search}%"))
    else:
        cur.execute("""
            SELECT id, filename, original_name, department, file_type,
                   file_size, uploader_name, uploader_id, created_at, updated_at
            FROM files WHERE status='Obsolete'
            ORDER BY updated_at DESC LIMIT 20
        """)
    result = []
    for r in cur.fetchall():
        f = dict_from_row(r)
        sz = f.get('file_size', 0) or 0
        if sz < 1024: f['size_fmt'] = f"{sz}B"
        elif sz < 1024*1024: f['size_fmt'] = f"{sz/1024:.1f}KB"
        else: f['size_fmt'] = f"{sz/1024/1024:.1f}MB"
        f['icon'] = file_ext_to_icon(f['original_name'])
        result.append(f)
    conn.close()
    return jsonify({'success': True, 'files': result})


@app.route('/api/obsolete/<file_id>', methods=['DELETE'])
@role_required('System Admin', 'Quality Specialist')
def api_delete_obsolete(file_id):
    """永久DeleteObsolete fileRecord及物理File"""
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute('SELECT * FROM files WHERE id=%s', (file_id,))
        row = cur.fetchone()
        if not row:
            cur.execute('SELECT * FROM obsolete_files WHERE id=%s', (file_id,))
            row = cur.fetchone()
        if not row:
            conn.close()
            return jsonify({'success': False, 'message': 'File Not存At'})
        # Will workflow_files.file_id （），Delete
        cur.execute('UPDATE workflow_files SET file_id=NULL WHERE file_id=%s', (file_id,))
        # DeleteFile
        fp = row.get('file_path')
        if not fp:
            fn = row.get('filename', '')
            for search_dir in [OBSOLETE_DIR, PUBLISHED_DIR]:
                for r, d, fs in os.walk(search_dir):
                    if fn in fs:
                        fp = os.path.join(r, fn)
                        break
        if fp and os.path.exists(fp):
            os.remove(fp)
            print(f'[ObsoleteDelete] 物理File Already Deleted: {fp}')
        # Delete files/obsolete_files Record（Approval file_id=NULL ForAlready Deleted）
        cur.execute('DELETE FROM files WHERE id=%s', (file_id,))
        cur.execute('DELETE FROM obsolete_files WHERE id=%s', (file_id,))
        conn.commit()
        conn.close()
        orig_name = row.get('original_name', file_id)
        log_operation(session['user_id'], session['username'], 'DeleteObsolete file', orig_name)
        return jsonify({'success': True, 'message': 'File Already永久Delete'})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'success': False, 'message': f'Delete Failed: {e}'}), 500


@app.route('/api/obsolete/register', methods=['POST'])
@role_required('System Admin', 'Quality Specialist')
def register_orphan_file():
    """Will孤立FileRegisterToData库"""
    file_path = request.form.get('file_path', '').strip()
    original_name = request.form.get('original_name', '').strip()
    department = request.form.get('department', '').strip()
    file_type = request.form.get('file_type', '').strip()
    reason = request.form.get('reason', 'Register孤立File').strip()
    
    if not file_path or not os.path.exists(file_path):
        return jsonify({'success': False, 'message': 'File Not存At'})
    if not original_name or not department or not file_type:
        return jsonify({'success': False, 'message': 'Please填写完整Information'})
    
    # CheckFileWhetherAlready AtData
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT id FROM files WHERE file_path=%s", (file_path,))
    if cur.fetchone():
        conn.close()
        return jsonify({'success': False, 'message': 'File already exists in the database'})
    
    # GetFileInformation
    filename = os.path.basename(file_path)
    file_size = os.path.getsize(file_path)
    file_id = str(uuid.uuid4())
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    # Data
    cur.execute("""
        INSERT INTO files (id, filename, original_name, department, file_type,
                           file_path, file_size, uploader_id, uploader_name, status,
                           is_published, created_at, updated_at, obsolete_reason)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
    """, (file_id, filename, original_name, department, file_type,
          file_path, file_size, session['user_id'], session['username'], 'Obsolete',
          0, now, now, reason))
    
    log_operation(session['user_id'], session['username'], 'Register孤立File',
                   f'{original_name} -> {department}/{file_type}')
    
    conn.commit()
    conn.close()
    
    return jsonify({'success': True, 'message': f'Already Registered: {original_name}'})



# =====================================================================
# User Management
# =====================================================================
@app.route('/users')
@role_required('System Admin')
def users():
    page = _int_param(request.args.get('page'), 1)
    per_page = 100
    offset = (page - 1) * per_page
    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT COUNT(*) AS cnt FROM users')
    total = cur.fetchone()['cnt']
    total_pages = max(1, (total + per_page - 1) // per_page)
    cur.execute('''
        SELECT id, username, email, department, cdsid, status, role, user_theme, must_change_password, created_at
        FROM users ORDER BY created_at DESC
        LIMIT %s OFFSET %s
    ''', (per_page, offset))
    user_list = [dict_from_row(r) for r in cur.fetchall()]
    conn.close()

    sidebar = get_sidebar_data(session['role'], session['user_id'])
    sidebar['user_theme'] = session.get('user_theme', 'tech-blue')
    return render_template('users.html', users=user_list, sidebar=sidebar, page_title='User Management',
                           page=page, total=total, total_pages=total_pages, per_page=per_page)

@app.route('/api/users/batch-delete', methods=['POST'])
@role_required('System Admin')
def batch_delete_users():
    data = request.get_json()
    ids = data.get('ids', [])
    if not ids:
        return jsonify({'success': False, 'message': '未SelectUser'})
    conn = get_db()
    try:
        cur = conn.cursor()
        placeholders = ','.join(['%s' for _ in ids])
        cur.execute(f"UPDATE users SET status='离职' WHERE id IN ({placeholders})", ids)
        conn.commit()
        affected = cur.rowcount
    except Exception as e:
        conn.rollback()
        raise
    finally:
        conn.close()
    log_operation(session['user_id'], session['username'], 'Batch Delete Users',
                   f'Batch set {len(ids)} users to resigned')
    return jsonify({'success': True, 'message': f'Successfully set {affected} users to resigned'})

@app.route('/api/users/permanent-delete', methods=['POST'])
@role_required('System Admin')
def permanent_delete_users():
    """Permanently delete users, clear foreign key references first"""
    data = request.get_json()
    ids = data.get('ids', [])
    if not ids:
        return jsonify({'success': False, 'message': 'No users selected'})
    
    # Prevent deleting yourself
    if session.get('user_id') in ids:
        return jsonify({'success': False, 'message': 'Cannot delete current logged-in user'})
    
    conn = get_db()
    try:
        cur = conn.cursor()
        placeholders = ','.join(['%s' for _ in ids])
        
        # 1. filesuploader_id
        cur.execute(f"UPDATE files SET uploader_id=NULL WHERE uploader_id IN ({placeholders})", ids)
        
        # 2. workflowsinitiator_id
        cur.execute(f"UPDATE workflows SET initiator_id=NULL WHERE initiator_id IN ({placeholders})", ids)
        
        # 3. workflow_recordsapprover_id
        cur.execute(f"UPDATE workflow_records SET approver_id=NULL WHERE approver_id IN ({placeholders})", ids)
        
        # 4. DeletenotificationsRelated records
        cur.execute(f"DELETE FROM notifications WHERE user_id IN ({placeholders})", ids)
        
        # 5. operation_logsuser_id（LogUser）
        cur.execute(f"UPDATE operation_logs SET user_id=NULL WHERE user_id IN ({placeholders})", ids)
        
        # 6. DeleteUserRecord
        cur.execute(f"DELETE FROM users WHERE id IN ({placeholders})", ids)
        
        conn.commit()
        affected = cur.rowcount
        del_count = len(ids)
    except Exception as e:
        conn.rollback()
        raise
    finally:
        conn.close()
    
    # Log after connection closed (avoid SQLite deadlock)
    log_operation(session['user_id'], session['username'], 'Permanent Delete Users',
                   f'Permanently deleted {del_count} users')
    
    return jsonify({'success': True, 'message': f'Successfully deleted {affected} users permanently'})


# ====== UserImportExport（5）======
import io

@app.route('/users/export')
@role_required('System Admin')
def export_users():
    """Export所UserFor Excel"""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT username, email, department, cdsid, role, status, user_theme, password_hash, created_at FROM users ORDER BY created_at')
    users = [dict_from_row(r) for r in cur.fetchall()]
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'User Information'

    #
    hdr_font = Font(name='微软雅黑', bold=True, color='FFFFFF', size=11)
    hdr_fill = PatternFill('solid', fgColor='1E3A5F')
    hdr_align = Alignment(horizontal='center', vertical='center')
    thin = Side(style='thin', color='CCCCCC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = ['Username', 'Email', 'Department', 'CDSID', 'Role', 'Status', '主题', 'Password哈希', 'Created At']
    for col, h in enumerate(headers, 1):
        cell = ws.cell(1, col, h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = hdr_align
        cell.border = border

    # Data
    data_font = Font(name='微软雅黑', size=10)
    for row_idx, u in enumerate(users, 2):
        vals = [
            u.get('username',''),
            u.get('email',''),
            u.get('department',''),
            u.get('cdsid',''),
            u.get('role',''),
            'Active' if u.get('status')=='Active' else '离职',
            u.get('user_theme','tech-blue'),
            u.get('password_hash',''),
            u.get('created_at','')
        ]
        for col, val in enumerate(vals, 1):
            cell = ws.cell(row_idx, col, val)
            cell.font = data_font
            cell.border = border
            cell.alignment = Alignment(horizontal='center', vertical='center')

    #
    widths = [18, 16, 28, 16, 18, 16, 16, 40, 10]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.row_dimensions[1].height = 30
    ws.freeze_panes = 'A2'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f'DMSUser Information_{datetime.now().strftime("%Y%m%d")}.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
@app.route('/users/import', methods=['GET'])
@role_required('System Admin')
def import_users_page():
    """Importpage面（仅Quality SpecialistCanOperation）"""
    sidebar = get_sidebar_data(session['role'], session['user_id'])
    sidebar['user_theme'] = session.get('user_theme', 'tech-blue')
    return render_template('user_import.html', sidebar=sidebar, page_title='Batch User Import')


@app.route('/users/import', methods=['POST'])
@role_required('System Admin')
def import_users():
    """Batch User Import"""
    import openpyxl
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': 'Please select Excel File'})

    file = request.files['file']
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'success': False, 'message': '仅Supports .xlsx / .xls File'})

    try:
        wb = openpyxl.load_workbook(file)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))

        if len(rows) < 2:
            return jsonify({'success': False, 'message': 'Excel File无效Data'})

        headers = [str(h).strip() if h else '' for h in rows[0]]
        #
        col_map = {}
        for idx, h in enumerate(headers):
            if 'Username' in h: col_map['username'] = idx
            elif 'Password' in h: col_map['password'] = idx
            elif 'Email' in h: col_map['email'] = idx
            elif 'Department' in h: col_map['department'] = idx
            elif 'CDSID' in h: col_map['cdsid'] = idx
            elif 'Role' in h: col_map['role'] = idx
            elif '主题' in h: col_map['user_theme'] = idx

        required = ['username', 'email']
        missing = [f for f in required if f not in col_map]
        if missing:
            return jsonify({'success': False, 'message': f'缺少必填列: {",".join(missing)}，Pleaseensure Excel 表头包含"Username"和"Email"'})

        conn = get_db()
        cur = conn.cursor()
        imported = 0
        errors = []

        for row_idx, row in enumerate(rows[1:], 2):
            try:
                username = str(row[col_map['username']]).strip() if col_map.get('username') is not None else ''
                password = str(row[col_map['password']]).strip() if col_map.get('password') is not None and col_map['password'] < len(row) and row[col_map['password']] is not None else ''
                email = str(row[col_map.get('email', 0)]).strip() if col_map.get('email') is not None and col_map['email'] < len(row) else ''
                department = str(row[col_map.get('department', 0)]).strip() if col_map.get('department') is not None and col_map['department'] < len(row) else ''
                cdsid = str(row[col_map.get('cdsid', 0)]).strip() if col_map.get('cdsid') is not None and col_map['cdsid'] < len(row) else ''
                role = str(row[col_map.get('role', 0)]).strip() if col_map.get('role') is not None and col_map['role'] < len(row) else 'Regular User'

                if not username or not email:
                    errors.append(f'Row {row_idx}: Username or email is empty')
                    continue

                # Check if already exists (case-insensitive)
                cur.execute('SELECT id FROM users WHERE LOWER(username)=LOWER(%s)', (username,))
                if cur.fetchone():
                    errors.append(f'Row {row_idx}: User "{username}" already exists, skipped')
                    continue

                user_id = str(uuid.uuid4())
                now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                user_theme = str(row[col_map.get('user_theme', 0)]).strip() if col_map.get('user_theme') is not None and col_map['user_theme'] < len(row) else 'tech-blue'
                # PasswordHandle：PasswordCan，ForPassword
                must_change = 0
                if password:
                    # Password: must be hash format (scrypt:/pbkdf2: or 64-char hex SHA256)
                    if password.startswith('scrypt:') or password.startswith('pbkdf2:'):
                        password_hash = password
                    elif len(password) == 64 and all(c in '0123456789abcdef' for c in password.lower()):
                        password_hash = password.lower()
                    else:
                        # Plain text password - not allowed
                        errors.append(f'Row {row_idx}: User "{username}" password must be hash format (scrypt:/pbkdf2:/SHA256), empty value will auto-generate random password')
                        continue
                else:
                    # PasswordFor：Password
                    random_pwd = secrets.token_urlsafe(6)[:8]
                    password_hash = generate_password_hash(random_pwd)
                    must_change = 1

                cur.execute('''
                    INSERT INTO users (id, username, password_hash, password, email, department, cdsid, role, status, user_theme, must_change_password, created_at)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                ''', (user_id, username, password_hash, '', email, department, cdsid, role, 'Active', user_theme, must_change, now))
                imported += 1

                # If auto-generated password, try to send email
                if not password and email:
                    try:
                        _send_password_reset_email(email, username, random_pwd)
                    except Exception:
                        pass  # Email send failure doesn't affect import
            except Exception as e:
                errors.append(f'Row {row_idx}: {str(e)}')

        conn.commit()
        conn.close()

        msg = f'Successfully imported {imported} users'
        if errors:
            msg += f', {len(errors)} skipped: {"; ".join(errors[:5])}'
            if len(errors) > 5:
                msg += f'... and {len(errors)} total'
        return jsonify({'success': True, 'message': msg, 'imported': imported, 'errors': errors[:10]})

    except Exception as e:
        return jsonify({'success': False, 'message': f'Failed to parse Excel: {str(e)}'})


@app.route('/api/users', methods=['POST'])
@role_required('System Admin')
def api_create_user():
    data = request.get_json()
    required = ['username', 'email']
    for f in required:
        if not data.get(f):
            return jsonify({'success': False, 'message': f'缺少必填field: {f}'})

    if len(data['username']) < 3:
        return jsonify({'success': False, 'message': 'Username Not少于3位'})

    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT id FROM users WHERE LOWER(username)=LOWER(%s)', (data['username'],))
    if cur.fetchone():
        conn.close()
        return jsonify({'success': False, 'message': 'Username already exists'})

    # Password
    random_pwd = secrets.token_urlsafe(6)[:8]
    user_id = str(uuid.uuid4())
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    cur.execute('''
        INSERT INTO users (id, username, password_hash, password, email, department, cdsid,
            status, role, user_theme, must_change_password, created_at)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
    ''', (user_id, data['username'],
          generate_password_hash(random_pwd), '',
          data.get('email', ''), data.get('department', ''),
          data.get('cdsid', ''), data.get('status', 'Active'),
          data.get('role', 'Regular User'), data.get('user_theme', 'tech-blue'),
          1, now))  # must_change_password=1
    conn.commit()
    conn.close()

    # PasswordReset
    email_addr = data.get('email', '')
    if email_addr:
        try:
            _send_password_reset_email(email_addr, data['username'], random_pwd)
        except Exception:
            pass  # Email send failedNotCreate

    log_operation(session['user_id'], session['username'], 'CreateUser',
                   f'CreateUser: {data["username"]}, Role: {data.get("role")}')
    return jsonify({'success': True, 'message': f'User Created Successfully，初始PasswordAlready发送至 {email_addr}'})

@app.route('/api/users/<user_id>', methods=['PUT'])
@role_required('System Admin')
def api_update_user(user_id):
    data = request.get_json()
    conn = get_db()
    cur = conn.cursor()
    fields = []
    params = []
    for k in ['email', 'department', 'cdsid', 'status', 'role', 'user_theme']:
        if k in data:
            fields.append(f'{k}=%s')
            params.append(data[k])
    # NotPassword
    if fields:
        params.append(user_id)
        cur.execute(f'UPDATE users SET {",".join(fields)} WHERE id=%s', params)
        conn.commit()
    conn.close()
    log_operation(session['user_id'], session['username'], 'UpdateUser',
                   f'UpdateUserID: {user_id}')
    return jsonify({'success': True, 'message': 'User Information Updated'})

@app.route('/api/users/<user_id>', methods=['DELETE'])
@role_required('System Admin')
def api_delete_user(user_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE users SET status='离职' WHERE id=%s", (user_id,))
    conn.commit()
    conn.close()
    log_operation(session['user_id'], session['username'], 'DeleteUser',
                   f'DeleteUserID: {user_id}')
    return jsonify({'success': True, 'message': 'User Already设For离职'})

@app.route('/api/change-password', methods=['POST'])
@login_required
def api_change_password():
    data = request.get_json()
    current_pwd = data.get('current_password', '')
    new_pwd = data.get('new_password', '')
    if not current_pwd or not new_pwd:
        return jsonify({'success': False, 'message': 'Please enter 前期 Password 和 新 Password'})
    if len(new_pwd) < 6:
        return jsonify({'success': False, 'message': '新PasswordNot少于6位'})
    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT id, password_hash FROM users WHERE id=%s', (session['user_id'],))
    user = cur.fetchone()
    if not user:
        conn.close()
        return jsonify({'success': False, 'message': 'User Not存At'})

    # PasswordVerify： werkzeug， SHA256 
    password_valid = False
    try:
        password_valid = check_password_hash(user['password_hash'], current_pwd)
    except Exception:
        pass
    if not password_valid:
        sha256_input = hashlib.sha256(current_pwd.encode()).hexdigest()
        password_valid = (user['password_hash'] == sha256_input)

    if not password_valid:
        conn.close()
        return jsonify({'success': False, 'message': 'When前PasswordNot正确'})

    # UpdatePassword：，Not
    cur.execute('UPDATE users SET password_hash=%s, password="", must_change_password=0 WHERE id=%s',
                (generate_password_hash(new_pwd), session['user_id']))
    conn.commit()
    conn.close()

    # session
    session.pop('must_change_password', None)

    return jsonify({'success': True, 'message': 'Password changed successfully'})

# =====================================================================
# Password / PasswordReset
# =====================================================================

def _send_password_reset_email(to_email, username, temp_password):
    """发送PasswordReset邮件"""
    smtp = get_smtp_config()
    if not smtp.get('enabled') or not smtp.get('host'):
        return False

    subject = 'DMS System - Password Reset Notification'
    body = f"""Hello, {username}:

Your password in the DMS Document Management System has been reset.

Temporary Password: {temp_password}

Please use this temporary password to log in to the system. After logging in, please change your password immediately.

If this is not your operation, please contact System Admin.

This email is automatically sent by the system. Please do not reply.
"""

    try:
        msg = mime_text.MIMEText(body, 'plain', 'utf-8')
        msg['Subject'] = email.header.Header(subject, 'utf-8')
        msg['From'] = smtp.get('from_addr', 'DMS@dms.local')
        msg['To'] = to_email
        port = int(smtp.get('port', 587))
        if port == 465:
            s = smtplib.SMTP_SSL(smtp['host'], port, timeout=8)
        else:
            s = smtplib.SMTP(smtp['host'], port, timeout=8)
            s.ehlo()
            if smtp.get('use_tls', True):
                s.starttls()
        if smtp.get('username') and smtp.get('password'):
            s.login(smtp['username'], smtp['password'])
        s.sendmail(msg['From'], [to_email], msg.as_string())
        s.quit()
        return True
    except Exception as e:
        print(f'[EMAIL] PasswordResetEmail send failed: {e}')
        return False


@app.route('/api/forgot-password', methods=['POST'])
def api_forgot_password():
    """忘记Password接口 - 根据EmailReset Password"""
    data = request.get_json() or {}
    email_addr = (data.get('email') or '').strip()

    if not email_addr:
        return jsonify({'success': False, 'message': 'Please enter Email Address'})

    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT id, username, email FROM users WHERE LOWER(email)=LOWER(%s) AND status="Active"',
                (email_addr,))
    user = cur.fetchone()

    if not user:
        conn.close()
        return jsonify({'success': False, 'message': '该Email未At系统中RegisterOrUser Already离职，Please联系Admin'})

    # 8AndPassword
    import string
    chars = string.ascii_letters + string.digits
    while True:
        temp_password = ''.join(secrets.choice(chars) for _ in range(8))
        has_letter = any(c.isalpha() for c in temp_password)
        has_digit = any(c.isdigit() for c in temp_password)
        if has_letter and has_digit:
            break

    # UpdatePasswordSettings
    cur.execute('UPDATE users SET password_hash=%s, password="", must_change_password=1 WHERE id=%s',
                (generate_password_hash(temp_password), user['id']))
    conn.commit()
    conn.close()

    #
    import threading
    email_sent = [False]
    def send_email_thread():
        email_sent[0] = _send_password_reset_email(email_addr, user['username'], temp_password)
    t = threading.Thread(target=send_email_thread, daemon=True)
    t.start()
    t.join(timeout=10)  # 10

    if email_sent[0]:
        return jsonify({'success': True, 'message': 'Reset Password邮件Already发送，Please查收Email'})
    else:
        return jsonify({'success': True, 'message': 'Password Already Reset，邮件发送中，Please稍后查收Email'})


# =====================================================================
# Approval
# =====================================================================

# =====================================================================
# Approval PDF Export
# =====================================================================
def _save_workflow_pdf(workflow_id, serial_no):
    """Generate PDF and save to workflow folder, return path or None if failed"""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib import colors
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    for fp in ['C:/Windows/Fonts/msyh.ttc', 'C:/Windows/Fonts/simhei.ttf', 'C:/Windows/Fonts/simsun.ttc']:
        if os.path.exists(fp):
            try:
                pdfmetrics.registerFont(TTFont('CNFont', fp)); cn = 'CNFont'; break
            except: pass
    else:
        cn = 'Helvetica'

    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT * FROM workflows WHERE id=%s', (workflow_id,))
    row = cur.fetchone()
    if not row:
        conn.close(); return None
    wf = dict_from_row(row)
    cur.execute('SELECT * FROM workflow_records WHERE workflow_id=%s ORDER BY created_at', (workflow_id,))
    records = [dict_from_row(r) for r in cur.fetchall()]
    cur.execute('SELECT * FROM workflow_files WHERE workflow_id=%s', (workflow_id,))
    files = [dict_from_row(r) for r in cur.fetchall()]
    conn.close()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=18*mm, rightMargin=18*mm, topMargin=15*mm, bottomMargin=15*mm)
    styles = getSampleStyleSheet()
    title_s = ParagraphStyle('T', parent=styles['Title'], fontName=cn, fontSize=14, leading=20)
    head_s = ParagraphStyle('H', parent=styles['Heading2'], fontName=cn, fontSize=11, leading=15)
    body_s = ParagraphStyle('B', parent=styles['Normal'], fontName=cn, fontSize=9, leading=13)
    elems = []
    elems.append(Paragraph('Approval Workflow Record - ' + str(wf.get('serial_no', '')), title_s))
    elems.append(Spacer(1, 5*mm))
    info_data = [
        [Paragraph('<b>Serial No.</b>', body_s), Paragraph(str(wf.get('serial_no', '')), body_s)],
        [Paragraph('<b>Title</b>', body_s), Paragraph(str(wf.get('title', '')), body_s)],
        [Paragraph('<b>Department</b>', body_s), Paragraph(str(wf.get('department', '')), body_s)],
        [Paragraph('<b>Initiator</b>', body_s), Paragraph(str(wf.get('initiator_name', '')), body_s)],
        [Paragraph('<b>Current Node</b>', body_s), Paragraph(str(wf.get('current_node', '')), body_s)],
        [Paragraph('<b>Status</b>', body_s), Paragraph(str(wf.get('status', '')), body_s)],
        [Paragraph('<b>Created At</b>', body_s), Paragraph(str(wf.get('created_at', '')), body_s)],
    ]
    t = Table(info_data, colWidths=[28*mm, 140*mm])
    t.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                           ('BACKGROUND', (0,0), (0,-1), colors.Color(0.9, 0.95, 1)),
                           ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                           ('LEFTPADDING', (0,0), (-1,-1), 5), ('TOPPADDING', (0,0), (-1,-1), 3), ('BOTTOMPADDING', (0,0), (-1,-1), 3)]))
    elems.append(t); elems.append(Spacer(1, 5*mm))
    if wf.get('apply_type') or wf.get('apply_reason') or wf.get('change_description'):
        elems.append(Paragraph('Application Information', head_s)); elems.append(Spacer(1, 2*mm))
        apply_data = []
        if wf.get('apply_type'):
            apply_data.append([Paragraph('<b>Application Type</b>', body_s), Paragraph(str(wf['apply_type']), body_s)])
        if wf.get('apply_reason'):
            apply_data.append([Paragraph('<b>Application Reason</b>', body_s), Paragraph(str(wf['apply_reason']), body_s)])
        if wf.get('change_description'):
            apply_data.append([Paragraph('<b>Change Description</b>', body_s), Paragraph(str(wf['change_description']), body_s)])
        if apply_data:
            at = Table(apply_data, colWidths=[28*mm, 140*mm])
            at.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                                    ('BACKGROUND', (0,0), (0,-1), colors.Color(0.9, 0.95, 1)),
                                    ('VALIGN', (0,0), (-1,-1), 'TOP'),
                                    ('LEFTPADDING', (0,0), (-1,-1), 5), ('TOPPADDING', (0,0), (-1,-1), 3), ('BOTTOMPADDING', (0,0), (-1,-1), 3)]))
            elems.append(at); elems.append(Spacer(1, 5*mm))
    elems.append(Paragraph('Approval Records', head_s)); elems.append(Spacer(1, 2*mm))
    if records:
        rd = [[Paragraph('<b>Role</b>', body_s), Paragraph('<b>Approver</b>', body_s), Paragraph('<b>Operation</b>', body_s), Paragraph('<b>Note</b>', body_s), Paragraph('<b>Time</b>', body_s)]]
        for r in records:
            rd.append([Paragraph(str(r.get('approver_role', '')), body_s), Paragraph(str(r.get('approver_name', '')), body_s), Paragraph(str(r.get('action', '')), body_s),
                      Paragraph(str(r.get('comment') or ''), body_s), Paragraph(str(r.get('created_at', '')), body_s)])
        rt = Table(rd, colWidths=[22*mm, 22*mm, 20*mm, 60*mm, 32*mm])
        rt.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.grey), ('BACKGROUND', (0,0), (-1,0), colors.Color(0.85, 0.9, 0.95)),
                                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LEFTPADDING', (0,0), (-1,-1), 4), ('TOPPADDING', (0,0), (-1,-1), 2), ('BOTTOMPADDING', (0,0), (-1,-1), 2)]))
        elems.append(rt)
    else:
        elems.append(Paragraph('No approval records', body_s))
    elems.append(Spacer(1, 5*mm)); elems.append(Paragraph('Attachments', head_s)); elems.append(Spacer(1, 2*mm))
    if files:
        for i, f in enumerate(files, 1):
            elems.append(Paragraph(str(i) + '. ' + str(f.get('original_name', '')) + ' (' + str(round(f.get('file_size', 0)/1024)) + 'KB)', body_s))
    else:
        elems.append(Paragraph('No attachments', body_s))
    doc.build(elems)
    buf.seek(0)
    pdf_bytes = buf.read()

    # SaveTo workflow/pdfs/ File
    pdf_dir = os.path.join(WORKFLOW_DIR, 'pdfs')
    os.makedirs(pdf_dir, exist_ok=True)
    safe_sn = serial_no.replace('/', '_').replace('\\', '_')
    out_path = os.path.join(pdf_dir, safe_sn + '_approval_records.pdf')
    with open(out_path, 'wb') as f:
        f.write(pdf_bytes)
    return out_path


@app.route('/api/workflow/<workflow_id>/pdf')
@login_required
def workflow_pdf(workflow_id):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib import colors
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    for fp in ['C:/Windows/Fonts/msyh.ttc', 'C:/Windows/Fonts/simhei.ttf', 'C:/Windows/Fonts/simsun.ttc']:
        if os.path.exists(fp):
            try:
                pdfmetrics.registerFont(TTFont('CNFont', fp))
                cn = 'CNFont'; break
            except: pass
    else:
        cn = 'Helvetica'

    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT * FROM workflows WHERE id=%s', (workflow_id,))
    wf = dict_from_row(cur.fetchone())
    if not wf:
        conn.close(); return jsonify({'success': False, 'message': '流程Not存At'}), 404
    cur.execute('SELECT * FROM workflow_records WHERE workflow_id=%s ORDER BY created_at', (workflow_id,))
    records = [dict_from_row(r) for r in cur.fetchall()]
    cur.execute('SELECT * FROM workflow_files WHERE workflow_id=%s', (workflow_id,))
    files = [dict_from_row(r) for r in cur.fetchall()]
    conn.close()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=18*mm, rightMargin=18*mm, topMargin=15*mm, bottomMargin=15*mm)
    styles = getSampleStyleSheet()
    title_s = ParagraphStyle('T', parent=styles['Title'], fontName=cn, fontSize=14, leading=20)
    head_s = ParagraphStyle('H', parent=styles['Heading2'], fontName=cn, fontSize=11, leading=15)
    body_s = ParagraphStyle('B', parent=styles['Normal'], fontName=cn, fontSize=9, leading=13)
    elems = []
    elems.append(Paragraph('Approval Workflow Record - ' + str(wf.get('serial_no', '')), title_s))
    elems.append(Spacer(1, 5*mm))
    info_data = [
        [Paragraph('<b>Serial No.</b>', body_s), Paragraph(str(wf.get('serial_no', '')), body_s)],
        [Paragraph('<b>Title</b>', body_s), Paragraph(str(wf.get('title', '')), body_s)],
        [Paragraph('<b>Department</b>', body_s), Paragraph(str(wf.get('department', '')), body_s)],
        [Paragraph('<b>Initiator</b>', body_s), Paragraph(str(wf.get('initiator_name', '')), body_s)],
        [Paragraph('<b>Current Node</b>', body_s), Paragraph(str(wf.get('current_node', '')), body_s)],
        [Paragraph('<b>Status</b>', body_s), Paragraph(str(wf.get('status', '')), body_s)],
        [Paragraph('<b>Created At</b>', body_s), Paragraph(str(wf.get('created_at', '')), body_s)],
    ]
    t = Table(info_data, colWidths=[28*mm, 140*mm])
    t.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                           ('BACKGROUND', (0,0), (0,-1), colors.Color(0.9, 0.95, 1)),
                           ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                           ('LEFTPADDING', (0,0), (-1,-1), 5), ('TOPPADDING', (0,0), (-1,-1), 3), ('BOTTOMPADDING', (0,0), (-1,-1), 3)]))
    elems.append(t); elems.append(Spacer(1, 5*mm))

    # PleaseInformation（field）
    if wf.get('apply_type') or wf.get('apply_reason') or wf.get('change_description'):
        elems.append(Paragraph('Application Information', head_s)); elems.append(Spacer(1, 2*mm))
        apply_data = []
        if wf.get('apply_type'):
            apply_data.append([Paragraph('<b>Application Type</b>', body_s), Paragraph(str(wf['apply_type']), body_s)])
        if wf.get('apply_reason'):
            apply_data.append([Paragraph('<b>Application Reason</b>', body_s), Paragraph(str(wf['apply_reason']), body_s)])
        if wf.get('change_description'):
            apply_data.append([Paragraph('<b>Change Description</b>', body_s), Paragraph(str(wf['change_description']), body_s)])
        if apply_data:
            at = Table(apply_data, colWidths=[28*mm, 140*mm])
            at.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                                    ('BACKGROUND', (0,0), (0,-1), colors.Color(0.9, 0.95, 1)),
                                    ('VALIGN', (0,0), (-1,-1), 'TOP'),
                                    ('LEFTPADDING', (0,0), (-1,-1), 5), ('TOPPADDING', (0,0), (-1,-1), 3), ('BOTTOMPADDING', (0,0), (-1,-1), 3)]))
            elems.append(at); elems.append(Spacer(1, 5*mm))

    elems.append(Paragraph('Approval Records', head_s)); elems.append(Spacer(1, 2*mm))
    if records:
        rd = [[Paragraph('<b>Role</b>', body_s), Paragraph('<b>Approver</b>', body_s), Paragraph('<b>Operation</b>', body_s), Paragraph('<b>Note</b>', body_s), Paragraph('<b>Time</b>', body_s)]]
        for r in records:
            rd.append([Paragraph(str(r.get('approver_role', '')), body_s), Paragraph(str(r.get('approver_name', '')), body_s), Paragraph(str(r.get('action', '')), body_s),
                      Paragraph(str(r.get('comment') or ''), body_s), Paragraph(str(r.get('created_at', '')), body_s)])
        rt = Table(rd, colWidths=[22*mm, 22*mm, 20*mm, 60*mm, 32*mm])
        rt.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.grey), ('BACKGROUND', (0,0), (-1,0), colors.Color(0.85, 0.9, 0.95)),
                                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LEFTPADDING', (0,0), (-1,-1), 4), ('TOPPADDING', (0,0), (-1,-1), 2), ('BOTTOMPADDING', (0,0), (-1,-1), 2)]))
        elems.append(rt)
    else:
        elems.append(Paragraph('No approval records', body_s))
    elems.append(Spacer(1, 5*mm)); elems.append(Paragraph('Attachments', head_s)); elems.append(Spacer(1, 2*mm))
    if files:
        for i, f in enumerate(files, 1): elems.append(Paragraph(str(i) + '. ' + str(f.get('original_name', '')) + ' (' + str(round(f.get('file_size', 0)/1024)) + 'KB)', body_s))
    else:
        elems.append(Paragraph('No attachments', body_s))
    doc.build(elems)
    buf.seek(0)
    fname = str(wf.get('serial_no', 'workflow')) + '_approval_records.pdf'
    # NotVersion Flask/Werkzeug
    try:
        return send_file(buf, as_attachment=True, download_name=fname, mimetype='application/pdf')
    except TypeError:
        # Version attachment_filename 
        return send_file(buf, as_attachment=True, attachment_filename=fname, mimetype='application/pdf')


@app.route('/approvals')
@login_required
def approvals():
    if session.get('role') == 'Read-only User':
        return redirect(url_for('dashboard'))
    role = session['role']
    user_id = session['user_id']
    conn = get_db()
    cur = conn.cursor()

    # UserApproval - In ProgressUserAtWhenApprovalList
    if role in ['System Admin', 'Quality Specialist']:
        # AdminQuality SpecialistCanToApproval
        cur.execute('''
            SELECT id, serial_no, title, department, initiator_name, current_node, status, created_at
            FROM workflows WHERE status='In Progress' AND is_closed=0
            ORDER BY created_at DESC LIMIT 500
        ''')
        pending_list = [dict_from_row(r) for r in cur.fetchall()]
    else:
        # Regular User：CanTo（1）NeedApproval + （2）InitiatedIn Progress
        cur.execute('''
            SELECT DISTINCT w.id, w.serial_no, w.title, w.department, w.initiator_name, w.current_node, w.status, w.created_at
            FROM workflows w
            LEFT JOIN workflow_nodes wn ON w.id = wn.workflow_id AND wn.node_name = w.current_node
            WHERE w.status='In Progress' AND w.is_closed=0
            AND (
                (wn.waiting_approvers LIKE %s)
                OR (w.initiator_id = %s)
            )
            ORDER BY w.created_at DESC LIMIT 500
        ''', (f'%{user_id}%', user_id))
        pending_list = [dict_from_row(r) for r in cur.fetchall()]

    # - Already（ApprovedOrReturned）
    if role in ['System Admin', 'Quality Specialist']:
        # AdminQuality SpecialistCanTo
        cur.execute('''
            SELECT id, serial_no, title, department, initiator_name, current_node, status, created_at
            FROM workflows 
            WHERE is_closed=1 OR status='Returned'
            ORDER BY created_at DESC LIMIT 500
        ''')
        history_list = [dict_from_row(r) for r in cur.fetchall()]
    else:
        # Regular User：CanToAndAlready
        # And：Initiated Or approval records
        cur.execute('''
            SELECT DISTINCT w.id, w.serial_no, w.title, w.department, w.initiator_name, w.current_node, w.status, w.created_at
            FROM workflows w
            WHERE (w.is_closed=1 OR w.status='Returned')
            AND (
                w.initiator_id = %s
                OR EXISTS (SELECT 1 FROM workflow_records wr WHERE wr.workflow_id = w.id AND wr.approver_id = %s)
            )
            ORDER BY w.created_at DESC LIMIT 500
        ''', (user_id, user_id))
        history_list = [dict_from_row(r) for r in cur.fetchall()]

    conn.close()

    sidebar = get_sidebar_data(role, user_id)
    sidebar['user_theme'] = session.get('user_theme', 'tech-blue')
    return render_template('approvals.html',
        pending_list=pending_list, history_list=history_list,
        sidebar=sidebar, page_title='File Approval',
        breadcrumb=[{'label': 'File Approval', 'url': None}])

def generate_serial_no():
    """生成流程Serial No.：YYYYMMDD + 四位No.（每日From0001开始）"""
    today = datetime.now().strftime('%Y%m%d')
    conn = get_db()
    cur = conn.cursor()
    # GetNo.
    cur.execute("SELECT MAX(serial_no) AS max_sn FROM workflows WHERE serial_no LIKE %s", (f'{today}%',))
    row = cur.fetchone()
    conn.close()
    if row and row['max_sn']:
        # No.+1
        last_seq = int(row['max_sn'][8:12])
        seq = last_seq + 1
    else:
        seq = 1
    return f'{today}{seq:04d}'


# Approval
APPROVAL_NODES = ['Quality Specialist', 'Dept Manager', 'Related Functional Manager', 'Quality Manager', 'Quality Release']
NODE_LABELS = {
    'Quality Specialist': 'Quality Specialist',
    'Dept Manager': 'Dept Manager',
    'Related Functional Manager': 'Related Functional Manager',
    'Quality Manager': 'Quality Manager',
    'Quality Release': 'Quality Release',
}


@app.route('/workflow/create', methods=['GET', 'POST'])
@login_required
def workflow_create():
    if session.get('role') == 'Read-only User':
        return redirect(url_for('dashboard'))
    if request.method == 'GET':
        conn = get_db()
        cur = conn.cursor()
        cur.execute('SELECT name FROM departments ORDER BY sort_order')
        depts = [r['name'] for r in cur.fetchall()]
        cur.execute('SELECT name FROM file_types ORDER BY sort_order')
        types = [r['name'] for r in cur.fetchall()]
        # ReturnUser InformationforSearch
        cur.execute("SELECT id, username, cdsid, department, role FROM users WHERE status='Active'")
        users = [dict_from_row(r) for r in cur.fetchall()]
        conn.close()

        sidebar = get_sidebar_data(session['role'], session['user_id'])
        sidebar['user_theme'] = session.get('user_theme', 'tech-blue')
        return render_template('workflow_create.html',
            departments=depts, file_types=types, users=users,
            sidebar=sidebar, page_title='Initiated Approval')

    # POST: SubmitApproval
    title = request.form.get('title', '').strip()
    if not title:
        return jsonify({'success': False, 'message': 'Please填写流程Title'})

    # Check5ApprovalWhetherAlready Selected
    node_data = {}  # {node_name: {ids: [], names: []}}
    for node_name in APPROVAL_NODES:
        node_name_key = node_name  # same as form key
        ids_str = request.form.get('node_0_ids' if node_name == 'Quality Specialist' else
                                   'node_1_ids' if node_name == 'Dept Manager' else
                                   'node_2_ids' if node_name == 'Related Functional Manager' else
                                   'node_3_ids' if node_name == 'Quality Manager' else 'node_4_ids', '')
        names_str = request.form.get('node_0_names' if node_name == 'Quality Specialist' else
                                      'node_1_names' if node_name == 'Dept Manager' else
                                      'node_2_names' if node_name == 'Related Functional Manager' else
                                      'node_3_names' if node_name == 'Quality Manager' else 'node_4_names', '')
        ids = [i.strip() for i in ids_str.split(',') if i.strip()]
        names = [n.strip() for n in names_str.split(',') if n.strip()]
        if not ids:
            return jsonify({'success': False, 'message': f'Please For「{node_name}」节点SelectApproval人'})
        node_data[node_name] = {'ids': ids, 'names': names}

    # Attachment
    file_count = int(request.form.get('file_count', '0'))
    apply_type = request.form.get('apply_type', '').strip()
    # ObsoleteType NotNeedUploadAttachment
    if file_count == 0 and apply_type != 'Obsolete':
        return jsonify({'success': False, 'message': 'Please至少Upload一个ApprovalAttachment'})

    # Serial No.
    serial_no = generate_serial_no()
    workflow_id = str(uuid.uuid4())
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # PleaseInformationfield（apply_type Already AtAttachmentGet）
    apply_reason = request.form.get('apply_reason', '').strip()
    change_description = request.form.get('change_description', '').strip()

    # Data
    conn = get_db()
    try:
        cur = conn.cursor()

        # GetNo.ApprovalForNo.
        first_node = APPROVAL_NODES[0]
        first_approver_ids = node_data[first_node]['ids']
        first_approver_names = node_data[first_node]['names']

        # AttachmentDepartmentForDepartment（No.AttachmentDepartment）
        wf_dept = ''
        wf_ftype = ''
        if file_count > 0:
            for i in range(file_count):
                dept_val = request.form.get(f'file_{i}_dept', '').strip()
                if dept_val:
                    wf_dept = dept_val
                    wf_ftype = request.form.get(f'file_{i}_type', '').strip()
                    break

        # CreateRecord
        cur.execute('''
            INSERT INTO workflows (id, serial_no, title, apply_type, apply_reason,
                change_description, initiator_id, initiator_name, department,
                current_node, status, created_at, updated_at)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
        ''', (workflow_id, serial_no, title, apply_type, apply_reason,
              change_description, session['user_id'], session['username'],
              wf_dept, first_node, 'In Progress', now, now))


        # 5Approval
        for idx, node_name in enumerate(APPROVAL_NODES):
            nd = node_data[node_name]
            can_multi = 1 if node_name == 'Related Functional Manager' else 0
            node_id = str(uuid.uuid4())
            cur.execute('''
                INSERT INTO workflow_nodes (id, workflow_id, node_name, node_order, is_required, can_multi, approver_ids, approver_names, waiting_approvers)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
            ''', (node_id, workflow_id, node_name, idx + 1, 1, can_multi,
                  ','.join(nd['ids']), ','.join(nd['names']), ','.join(nd['ids'])))

        # Attachment（AttachmentDepartment+Type）
        wf_dir = os.path.join(WORKFLOW_DIR, workflow_id)
        os.makedirs(wf_dir, exist_ok=True)
        for i in range(file_count):
            f = request.files.get(f'file_{i}')
            if not f or not f.filename:
                continue
            dept = request.form.get(f'file_{i}_dept', '').strip()
            ftype = request.form.get(f'file_{i}_type', '').strip()
            original_name = f.filename
            ext = os.path.splitext(original_name)[1]
            saved_name = f'{uuid.uuid4().hex}{ext}'
            file_path = os.path.join(wf_dir, saved_name)
            f.save(file_path)
            file_size = os.path.getsize(file_path)
            wf_file_id = str(uuid.uuid4())
            cur.execute('''
                INSERT INTO workflow_files (id, workflow_id, filename, original_name,
                    department, file_type, file_path, file_size, created_at)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
            ''', (wf_file_id, workflow_id, saved_name, original_name,
                  dept, ftype, file_path, file_size, now))

        # InitiatedRecord
        record_id = str(uuid.uuid4())
        cur.execute('''
            INSERT INTO workflow_records (id, workflow_id, node_name, approver_id,
                approver_name, approver_role, `action`, `comment`, created_at)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
        ''', (record_id, workflow_id, 'Initiated', session['user_id'], session['username'],
              'Initiator', 'Initiated', '', now))

        conn.commit()

        # NotificationNo.Approval
        for i, uid in enumerate(first_approver_ids):
            uname = first_approver_names[i] if i < len(first_approver_names) else ''
            uemail = _get_user_email(cur, uid)
            notify_workflow(workflow_id, 'Initiated', first_node, uid, uname, uemail, '')

        serial_no_ref = serial_no
        title_ref = title
    except Exception as e:
        conn.rollback()
        raise
    finally:
        conn.close()
    
    # AtRecordLog（avoidSQLite）
    log_operation(session['user_id'], session['username'], 'Initiated Approval',
                 f'Initiated Approval: {serial_no_ref} {title_ref}')
    return jsonify({'success': True, 'message': f'Approval流程AlreadyInitiated，Serial No.: {serial_no_ref}', 'workflow_id': workflow_id})

@app.route('/workflow/<workflow_id>/download/<file_id>')
@login_required
def download_workflow_file(workflow_id, file_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT * FROM workflow_files WHERE id=%s AND workflow_id=%s', (file_id, workflow_id))
    wf = cur.fetchone()
    if not wf:
        conn.close()
        return render_template('error.html', message='File Not存At')
    wf = dict_from_row(wf)
    # Published file_id=NULL Already ObsoleteDelete，NotDownload
    if wf.get('is_published') and wf.get('file_id') is None:
        conn.close()
        return render_template('error.html',
            message=f'Attachment "{wf["original_name"]}" Already被Obsolete并彻底Delete，无法Download')
    conn.close()
    directory = os.path.dirname(wf['file_path'])
    filename = wf['filename']
    original_name = wf['original_name']
    log_operation(session['user_id'], session['username'], 'Download approvalAttachment',
                  f'{original_name} [{workflow_id[:8]}]')
    return send_from_directory(directory, filename,
                               as_attachment=True,
                               download_name=original_name)

@app.route('/workflow/<workflow_id>')
@login_required
def workflow_detail(workflow_id):
    if session.get('role') == 'Read-only User':
        return redirect(url_for('dashboard'))
    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT * FROM workflows WHERE id=%s', (workflow_id,))
    wf = cur.fetchone()
    if not wf:
        conn.close()
        return render_template('error.html', message='Approval流程Not存At')
    wf = dict_from_row(wf)

    # GetStatus（5 + FixReturnedStatus）
    cur.execute('SELECT * FROM workflow_nodes WHERE workflow_id=%s ORDER BY node_order', (workflow_id,))
    nodes = [dict_from_row(r) for r in cur.fetchall()]

    current_node = wf['current_node']

    # Get approval records（Check/ReturnedStatus）
    cur.execute('SELECT node_name, `action` FROM workflow_records WHERE workflow_id=%s ORDER BY created_at',
               (workflow_id,))
    all_records = [dict_from_row(r) for r in cur.fetchall()]

    # Already
    done_nodes = set()
    rejected_nodes = set()
    for rec in all_records:
        if rec['action'] == 'Approved':
            done_nodes.add(rec['node_name'])
        elif rec['action'] == 'Returned':
            rejected_nodes.add(rec['node_name'])

    for node in nodes:
        nid = node['node_name']
        if nid == current_node and current_node not in ('Initiated', '待发布'):
            node['node_status'] = 'current'
        elif nid in rejected_nodes:
            node['node_status'] = 'rejected'
        elif nid in done_nodes:
            node['node_status'] = 'done'
        else:
            node['node_status'] = 'pending'

        # ApprovalName（2：Related Functional Manager）
        if nid == current_node and node.get('waiting_approvers'):
            waiting_ids = [i.strip() for i in node['waiting_approvers'].split(',') if i.strip()]
            if waiting_ids:
                placeholders = ','.join(['%s'] * len(waiting_ids))
                cur.execute(f'SELECT username FROM users WHERE id IN ({placeholders})', tuple(waiting_ids))
                node['waiting_names'] = '、'.join([r['username'] for r in cur.fetchall()])

    # Get approval records（7）
    cur.execute('''
        SELECT * FROM workflow_records WHERE workflow_id=%s ORDER BY created_at
    ''', (workflow_id,))
    records = [dict_from_row(r) for r in cur.fetchall()]

    # GetAttachment
    cur.execute('''
        SELECT * FROM workflow_files WHERE workflow_id=%s ORDER BY created_at
    ''', (workflow_id,))
    files = [dict_from_row(r) for r in cur.fetchall()]
    for f in files:
        size = f.get('file_size', 0) or 0
        f['size_fmt'] = f'{size/1024:.1f}KB' if size < 1024*1024 else f'{size/1024/1024:.1f}MB'
        f['icon'] = file_ext_to_icon(f['original_name'])
        # Already Deleted（FileDelete）
        f['is_deleted'] = bool(f.get('is_deleted'))

    conn.close()

    # CheckPermission
    user_id = session['user_id']
    role = session['role']

    # WhenApprovalIDs
    current_node_obj = next((n for n in nodes if n['node_name'] == current_node), None)
    current_approver_ids = []
    if current_node_obj and current_node_obj['approver_ids']:
        current_approver_ids = [i.strip() for i in current_node_obj['approver_ids'].split(',') if i.strip()]

    # InitiatedWhetherAlreadyFromInitiatedSubmit（forReturnedSubmit）
    has_submitted = (
        current_node != 'Initiated' and
        wf['initiator_id'] == user_id and
        any(r['action'] == '重新Submit' for r in records)
    )

    # WhenUserWhetherAlready AtWhenApproval（At can_approve ）
    has_approved = (
        any(r['approver_id'] == user_id and r['node_name'] == current_node
            for r in records)
    )

    # WhetherCanWhenUserApproval
    can_approve = (
        wf['status'] == 'In Progress' and
        current_node != 'Initiated' and
        user_id in current_approver_ids and
        not has_approved
    )
    
    # Permission：Quality Specialist + Approved + + ObsoleteType
    can_publish = (
        role == 'Quality Specialist' and
        wf['status'] == 'Approved' and
        not wf.get('published') and
        wf.get('apply_type') != 'Obsolete'
    )
    is_published = bool(wf.get('published'))


    # Department+TypeList
    conn2 = get_db()
    cur2 = conn2.cursor()
    cur2.execute('SELECT name FROM departments ORDER BY sort_order')
    departments = [r['name'] for r in cur2.fetchall()]
    cur2.execute('SELECT name FROM file_types ORDER BY sort_order')
    file_types = [r['name'] for r in cur2.fetchall()]
    conn2.close()

    sidebar = get_sidebar_data(role, user_id)
    sidebar['user_theme'] = session.get('user_theme', 'tech-blue')
    return render_template('workflow_detail.html',
        workflow=wf, records=records, workflow_files=files,
        workflow_nodes=nodes,
        can_approve=can_approve,
        has_approved=has_approved,
        can_publish=can_publish,
        is_published=is_published,
        departments=departments, file_types=file_types,
        sidebar=sidebar, page_title=f'Approval Workflow {wf["serial_no"]}')

@app.route('/workflow/<workflow_id>/process', methods=['POST'])
@login_required
def workflow_process(workflow_id):
    data = request.get_json()
    action = data.get('action', '')
    comment = data.get('comment', '') or ''

    if action not in ('Approved', 'Returned'):
        return jsonify({'success': False, 'message': 'Unknown operation'})

    if len(comment) > 1000:
        return jsonify({'success': False, 'message': 'Note cannot exceed 1000 characters'})

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    user_id = session['user_id']

    conn = get_db()
    try:
        cur = conn.cursor()

        cur.execute('SELECT * FROM workflows WHERE id=%s', (workflow_id,))
        wf = cur.fetchone()
        if not wf:
            return jsonify({'success': False, 'message': 'Workflow not found'})

        wf = dict_from_row(wf)

        if wf['status'] != 'In Progress':
            return jsonify({'success': False, 'message': 'Workflow already ended'})

        current_node = wf['current_node']

        cur.execute('SELECT approver_ids, can_multi, waiting_approvers FROM workflow_nodes WHERE workflow_id=%s AND node_name=%s',
                   (workflow_id, current_node))
        node_row = cur.fetchone()
        if not node_row or not node_row['approver_ids']:
            return jsonify({'success': False, 'message': 'Current node has no approver'})

        approver_ids_all = [i.strip() for i in node_row['approver_ids'].split(',') if i.strip()]
        can_multi = node_row['can_multi'] or 0
        waiting_str = node_row['waiting_approvers'] or ','.join(approver_ids_all)
        waiting_list = [i.strip() for i in waiting_str.split(',') if i.strip()]

        if user_id not in approver_ids_all:
            return jsonify({'success': False, 'message': 'You are not the approver of current node'})

        record_id = str(uuid.uuid4())
        cur.execute('''
            INSERT INTO workflow_records (id, workflow_id, node_name, approver_id,
                approver_name, approver_role, `action`, `comment`, created_at)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
        ''', (record_id, workflow_id, current_node, user_id,
              session['username'], current_node, action, comment, now))

        if action == 'Returned':
            cur.execute("UPDATE workflows SET status='Returned', current_node=%s, updated_at=%s WHERE id=%s",
                       (current_node, now, workflow_id))
            conn.commit()
            # Notification Initiated
            notify_workflow(workflow_id, 'Returned', 'Returned', wf['initiator_id'],
                          wf['initiator_name'], _get_user_email(cur, wf['initiator_id']), comment)
            msg = 'Workflow returned, process terminated'
            log_operation(user_id, session['username'], 'Approval-Returned',
                         f'Workflow {wf["serial_no"]}: {current_node} returned, note: {comment or "none"}')
            return jsonify({'success': True, 'message': msg, 'status': 'Returned'})

        # Approved
        if user_id in waiting_list:
            waiting_list.remove(user_id)

        if can_multi and waiting_list:
            cur.execute("UPDATE workflow_nodes SET waiting_approvers=%s WHERE workflow_id=%s AND node_name=%s",
                       (','.join(waiting_list), workflow_id, current_node))
            conn.commit()
            pending_names = _get_names_from_ids(cur, waiting_list)
            msg = f'You have approved, waiting for others: {pending_names}'
            log_operation(user_id, session['username'], 'Approval-Approved',
                         f'Workflow {wf["serial_no"]}: {current_node} approved (waiting: {pending_names})')
            return jsonify({'success': True, 'message': msg, 'waiting': pending_names})

        # All approvers at this node have approved, move to next node or end
        next_node = get_next_node(current_node)

        if next_node:
            # Move to next node
            cur.execute("UPDATE workflows SET current_node=%s, updated_at=%s WHERE id=%s",
                       (next_node, now, workflow_id))
            conn.commit()
            _notify_next_node(cur, workflow_id, next_node, comment)
            msg = f'Approved, current node: {next_node}'
            log_operation(user_id, session['username'], 'Approval-Approved',
                         f'Workflow {wf["serial_no"]}: {current_node} -> {next_node}')
            return jsonify({'success': True, 'message': msg})
        else:
            # Last node, workflow completed
            cur.execute("UPDATE workflows SET status='Approved', current_node=%s, is_closed=1, closed_at=%s, updated_at=%s WHERE id=%s",
                       (current_node, now, now, workflow_id))
            conn.commit()
            notify_workflow(workflow_id, 'Completed', 'Approved', wf['initiator_id'],
                          wf['initiator_name'], _get_user_email(cur, wf['initiator_id']), comment)
            msg = 'Approved, workflow completed'
            log_operation(user_id, session['username'], 'Approved',
                         f'Workflow {wf["serial_no"]}: all nodes passed')
            
            # Auto-generate PDF when approved
            try:
                _save_workflow_pdf(workflow_id, wf['serial_no'])
            except Exception as e:
                print(f'[PDF Save Failed] Workflow {wf["serial_no"]}: {e}')
            
            return jsonify({'success': True, 'message': msg, 'status': 'Approved'})
    except Exception as e:
        conn.rollback()
        raise
    finally:
        conn.close()


def _get_names_from_ids(cur, ids):
    if not ids:
        return ''
    placeholders = ','.join(['%s'] * len(ids))
    cur.execute(f'SELECT username FROM users WHERE id IN ({placeholders})', tuple(ids))
    rows = cur.fetchall()
    return '、'.join([r['username'] for r in rows])


@app.route('/workflow/<workflow_id>/cancel', methods=['POST'])
@login_required
def workflow_cancel(workflow_id):
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT * FROM workflows WHERE id=%s', (workflow_id,))
    wf = cur.fetchone()
    if not wf:
        conn.close()
        return jsonify({'success': False, 'message': '流程Not存At'})
    wf = dict_from_row(wf)

    if wf['initiator_id'] != session['user_id']:
        conn.close()
        return jsonify({'success': False, 'message': '只Initiated人Can以Cancel'})

    if wf['status'] != 'In Progress':
        conn.close()
        return jsonify({'success': False, 'message': '流程Already结束'})

    cur.execute("UPDATE workflows SET status='Already Cancelled', current_node='Already Cancelled', is_closed=1, closed_at=%s WHERE id=%s",
               (now, workflow_id))
    # RecordCancel
    record_id = str(uuid.uuid4())
    cur.execute('''
        INSERT INTO workflow_records (id, workflow_id, node_name, approver_id,
            approver_name, approver_role, `action`, `comment`, created_at)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
    ''', (record_id, workflow_id, 'Initiated', session['user_id'], session['username'],
          'Initiator', 'Cancel', '', now))
    conn.commit()
    conn.close()
    log_operation(session['user_id'], session['username'], 'Cancel approval',
                 f'流程 {wf["serial_no"]}')
    return jsonify({'success': True, 'message': 'Approval Already Cancelled'})


# ====== Can（） =====
@app.route('/workflow/<workflow_id>/publish', methods=['POST'])
@role_required('Quality Specialist', 'System Admin')
def workflow_publish(workflow_id):
    """WillApproval流程中Attachment发布ToFile库"""
    import shutil
    
    conn = get_db()
    cur = conn.cursor()
    
    # CheckWhetherAtCompleted
    cur.execute('SELECT * FROM workflows WHERE id=%s', (workflow_id,))
    wf = cur.fetchone()
    if not wf:
        conn.close()
        return jsonify({'success': False, 'message': '流程Not存At'})
    wf = dict_from_row(wf)
    
    # CheckStatus：Approved
    if wf['status'] != 'Approved':
        conn.close()
        return jsonify({'success': False, 'message': '只Approved流程才Can发布'})
    
    # CheckWhetherPublished
    if wf.get('published'):
        conn.close()
        return jsonify({'success': False, 'message': '此流程中Attachment Published，Cannot重复发布'})
    
    # GetAttachment
    cur.execute('SELECT * FROM workflow_files WHERE workflow_id=%s', (workflow_id,))
    wf_files = cur.fetchall()
    
    if not wf_files:
        conn.close()
        return jsonify({'success': False, 'message': '此流程Attachment'})
    
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    published_count = 0
    publish_details = []  # Details，AtRecordLog
    published_files = []  # [(file_id, target_path, original_name)] forIndexUpdate
    
    for wf_file in wf_files:
        wf_file = dict_from_row(wf_file)
        
        # PublishedAttachment
        if wf_file.get('is_published'):
            continue
        
        # File Path
        src_path = wf_file['file_path']
        if not src_path or not os.path.exists(src_path):
            continue
        
        # Directory：PUBLISHED_DIR/Department/File Type/
        dept = wf_file['department']
        file_type = wf_file['file_type']
        if not dept or not file_type:
            continue
        
        target_dir = os.path.join(PUBLISHED_DIR, dept, file_type)
        os.makedirs(target_dir, exist_ok=True)
        
        # File Name（avoid）
        ext = os.path.splitext(wf_file['original_name'])[1]
        new_filename = f'{uuid.uuid4().hex}{ext}'
        target_path = os.path.join(target_dir, new_filename)
        
        # FileToDirectory
        shutil.copy2(src_path, target_path)
        
        # files 
        file_id = str(uuid.uuid4())
        cur.execute('''
            INSERT INTO files (id, filename, original_name, department, file_type,
                file_path, file_size, status, uploader_id, uploader_name,
                created_at, updated_at, is_published)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,1)
        ''', (file_id, new_filename, wf_file['original_name'], dept, file_type,
              target_path, wf_file['file_size'], 'Published', session['user_id'],
              session['username'], now, now))
        
        # Update workflow_files ForPublished
        cur.execute('UPDATE workflow_files SET is_published=1, file_id=%s WHERE id=%s',
                   (file_id, wf_file['id']))
        
        published_count += 1
        publish_details.append(f'{wf_file["original_name"]} -> {dept}/{file_type}')
        published_files.append((new_filename, target_path, wf_file['original_name']))
    
    # Published
    cur.execute('UPDATE workflows SET published=1 WHERE id=%s', (workflow_id,))
    conn.commit()
    conn.close()
    
    # AtRecordLog（avoidSQLite）
    if published_count == 0:
        return jsonify({'success': False, 'message': 'Can发布Attachment（CanCan全部PublishedOr源File丢失）'})
    
    for detail in publish_details:
        log_operation(session['user_id'], session['username'], '发布File', detail)
    log_operation(session['user_id'], session['username'], '一键发布',
                 f'流程 {wf["serial_no"]}，发布 {published_count} 个File')

    # Full-text indexUpdate（，NotReturn）
    import threading
    for fid, fpath, fname in published_files:
        t = threading.Thread(target=_update_file_index, args=(fpath, fname, fid), daemon=True)
        t.start()

    return jsonify({'success': True, 'message': f'发布Success， Total 发布 {published_count} 个FileToFile库'})


# ====== 4: ReturnedInitiatedAttachment ======
@app.route('/workflow/<workflow_id>/add-file', methods=['POST'])
@login_required
def workflow_add_file(workflow_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT * FROM workflows WHERE id=%s', (workflow_id,))
    wf = cur.fetchone()
    if not wf:
        conn.close()
        return jsonify({'success': False, 'message': '流程Not存At'})
    wf = dict_from_row(wf)

    if wf['initiator_id'] != session['user_id']:
        conn.close()
        return jsonify({'success': False, 'message': '只Initiated人Can以Operation'})

    if wf['status'] != 'In Progress' or wf['current_node'] != 'Initiated':
        conn.close()
        return jsonify({'success': False, 'message': 'When前StatusNot允许添加Attachment'})

    f = request.files.get('file')
    if not f or not f.filename:
        conn.close()
        return jsonify({'success': False, 'message': 'Please select File'})

    dept = request.form.get('department', '').strip()
    ftype = request.form.get('file_type', '').strip()
    if not dept or not ftype:
        conn.close()
        return jsonify({'success': False, 'message': 'Please select Department and File Type'})

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    original_name = f.filename
    ext = os.path.splitext(original_name)[1]
    saved_name = f'{uuid.uuid4().hex}{ext}'
    wf_dir = os.path.join(WORKFLOW_DIR, workflow_id)
    os.makedirs(wf_dir, exist_ok=True)
    file_path = os.path.join(wf_dir, saved_name)
    f.save(file_path)
    file_size = os.path.getsize(file_path)
    wf_file_id = str(uuid.uuid4())
    cur.execute('''
        INSERT INTO workflow_files (id, workflow_id, filename, original_name,
            department, file_type, file_path, file_size, created_at)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
    ''', (wf_file_id, workflow_id, saved_name, original_name,
          dept, ftype, file_path, file_size, now))
    conn.commit()
    conn.close()
    log_operation(session['user_id'], session['username'], '添加ApprovalAttachment', original_name)
    return jsonify({'success': True, 'message': 'Attachment Already添加'})


@app.route('/workflow/<workflow_id>/remove-file/<file_id>', methods=['POST'])
@login_required
def workflow_remove_file(workflow_id, file_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT * FROM workflows WHERE id=%s', (workflow_id,))
    wf = cur.fetchone()
    if not wf:
        conn.close()
        return jsonify({'success': False, 'message': '流程Not存At'})
    wf = dict_from_row(wf)

    if wf['initiator_id'] != session['user_id']:
        conn.close()
        return jsonify({'success': False, 'message': '只Initiated人Can以DeleteAttachment'})

    if wf['status'] != 'In Progress' or wf['current_node'] != 'Initiated':
        conn.close()
        return jsonify({'success': False, 'message': 'When前StatusNot允许DeleteAttachment'})

    cur.execute('DELETE FROM workflow_files WHERE id=%s AND workflow_id=%s', (file_id, workflow_id))
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': 'Attachment Already Deleted'})


# ====== 6: ======
import zipfile

WORKFLOW_ARCHIVES_DIR = os.path.join(UPLOAD_DIR, 'workflow_archives')
os.makedirs(WORKFLOW_ARCHIVES_DIR, exist_ok=True)


@app.route('/workflow/<workflow_id>/delete', methods=['POST'])
@login_required
def workflow_delete(workflow_id):
    """AdminDelete approval流程"""
    if session['role'] != 'System Admin':
        return jsonify({'success': False, 'message': 'Permission Not足，仅System AdminCanDelete approval流程'})

    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT * FROM workflows WHERE id=%s', (workflow_id,))
    wf = cur.fetchone()
    if not wf:
        conn.close()
        return jsonify({'success': False, 'message': '流程Not存At'})

    wf = dict_from_row(wf)

    # DeleteFile
    wf_dir = os.path.join(WORKFLOW_DIR, workflow_id)
    if os.path.exists(wf_dir):
        import shutil
        shutil.rmtree(wf_dir)

    # Delete approvalRecordPDF
    pdf_dir = os.path.join(WORKFLOW_DIR, 'pdfs')
    pdf_path = os.path.join(pdf_dir, wf['serial_no'] + '_approval_records.pdf')
    if os.path.exists(pdf_path):
        os.remove(pdf_path)

    # DeleteDataRecord（Attachment、Record、，Record）
    cur.execute('DELETE FROM workflow_files WHERE workflow_id=%s', (workflow_id,))
    cur.execute('DELETE FROM workflow_records WHERE workflow_id=%s', (workflow_id,))
    cur.execute('DELETE FROM workflow_nodes WHERE workflow_id=%s', (workflow_id,))
    cur.execute('DELETE FROM notifications WHERE workflow_id=%s', (workflow_id,))
    cur.execute('DELETE FROM workflows WHERE id=%s', (workflow_id,))
    conn.commit()
    conn.close()

    log_operation(session['user_id'], session['username'], 'Delete approval流程',
                 f'Delete流程 {wf["serial_no"]} {wf["title"]}')
    return jsonify({'success': True, 'message': f'流程 {wf["serial_no"]} Already Deleted'})


@app.route('/api/workflows/search')
@login_required
def api_workflows_search():
    """SearchApproval流程（按流程号、Initiated人、Attachment File Name）"""
    q = request.args.get('q', '').strip()
    if not q:
        return jsonify({'results': [], 'count': 0})

    role = session.get('role', '')
    user_id = session.get('user_id')
    conn = get_db()
    cur = conn.cursor()

    p = f'%{q}%'

    # Search：Search、Title、Initiated，Attachment File Name
    # DISTINCT （CanCanAttachment）
    query = '''
        SELECT DISTINCT w.id, w.serial_no, w.title, w.department, w.initiator_name,
                        w.current_node, w.status, w.created_at, w.is_closed
        FROM workflows w
        LEFT JOIN workflow_files wf ON wf.workflow_id = w.id
        WHERE (
            w.serial_no LIKE %s OR
            w.title LIKE %s OR
            w.initiator_name LIKE %s OR
            wf.original_name LIKE %s
        )
    '''
    params = [p, p, p, p]

    # Permission：Regular UserCanInitiatedOrAnd
    if role not in ['System Admin', 'Quality Specialist']:
        query += ''' AND (
            w.initiator_id = %s OR
            w.id IN (
                SELECT DISTINCT workflow_id FROM workflow_records WHERE approver_id = %s
            )
        )'''
        params.extend([user_id, user_id])

    query += ' ORDER BY w.created_at DESC LIMIT 50'
    cur.execute(query, params)
    rows = cur.fetchall()

    results = []
    for r in rows:
        wf = dict_from_row(r)
        # GetAttachments（forFile Name）
        cur.execute('SELECT original_name FROM workflow_files WHERE workflow_id=%s', (wf['id'],))
        files = [f['original_name'] for f in cur.fetchall()]
        wf['attachments'] = files
        wf['attachment_count'] = len(files)
        #
        wf['highlight_serial'] = wf['serial_no'].replace(q, f'<mark style="background:rgba(0,212,255,0.3);color:inherit;border-radius:2px;padding:0 2px;">{q}</mark>')
        wf['highlight_title'] = wf['title'].replace(q, f'<mark style="background:rgba(0,212,255,0.3);color:inherit;border-radius:2px;padding:0 2px;">{q}</mark>')
        results.append(wf)

    conn.close()
    return jsonify({'results': results, 'count': len(results)})


def _get_user_email(cur, user_id):
    """GetUserEmail"""
    cur.execute('SELECT email FROM users WHERE id=%s', (user_id,))
    row = cur.fetchone()
    return row['email'] if row else ''


def _notify_next_node(cur, workflow_id, next_node, comment=''):
    """NotificationNext NodeApproval人"""
    # From workflow_nodes Next NodeApproval
    cur.execute('''
        SELECT approver_ids FROM workflow_nodes
        WHERE workflow_id=%s AND node_name=%s LIMIT 1
    ''', (workflow_id, next_node))
    row = cur.fetchone()
    print(f'[NOTIFY] _notify_next_node: workflow={workflow_id[:8]}, next_node={next_node}, row={row}')
    if row and row['approver_ids']:
        approver_ids = row['approver_ids'].split(',')
        for uid in approver_ids:
            uid = uid.strip()
            if not uid:
                continue
            cur.execute('SELECT username, email FROM users WHERE id=%s', (uid,))
            u = cur.fetchone()
            print(f'[NOTIFY] Approver: uid={uid}, user={u}')
            if u:
                notify_workflow(workflow_id, 'Approved', next_node, uid,
                              u['username'], u['email'], comment)


def get_next_node(current):
    """按顺序GetNext Node"""
    nodes = APPROVAL_NODES  # ['Quality Specialist', 'Dept Manager', 'Related Functional Manager', 'Quality Manager', 'Quality Release']
    try:
        idx = nodes.index(current)
        return nodes[idx+1] if idx+1 < len(nodes) else None
    except ValueError:
        return None

def get_smtp_config():
    """FromConfig读取 SMTP Settings"""
    try:
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            cfg = _json.load(f)
        return cfg.get('smtp', {})
    except Exception:
        return {}

def get_scheduler_config():
    """FromConfig读取定时任务Settings，默认关闭"""
    try:
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            cfg = _json.load(f)
        return cfg.get('scheduler', {'weekly_index_enabled': False, 'weekly_cleanup_enabled': False})
    except Exception:
        return {'weekly_index_enabled': False, 'weekly_cleanup_enabled': False}

def _send_email_background(to_email, to_name, serial_no, title, action, next_node, workflow_id, comment, attachment_name):
    """后台线程发送邮件（异步，Not阻塞Approval响Should）"""
    smtp = get_smtp_config()
    if not smtp.get('enabled') or not smtp.get('host'):
        return

    action_texts = {
        'Initiated': 'You have a new approval task pending.',
        'Approved': f'You have a new approval task pending. Current node: {next_node}',
        'Returned': 'The approval has been returned. Please note.',
        'Completed': 'Approval workflow completed. All attachments have been published.',
    }
    if action == 'Initiated' or action == 'Approved':
        subject = f'Please review the approval: [{serial_no}][{attachment_name}]'
    else:
        subject = f'[{serial_no}] {title} - {action_texts.get(action, action)}'
    action_desc = action_texts.get(action, action)

    body = f"""Hello, {to_name or to_email}:

{action_desc}

━━━━━━━━━━━━━━━━━━━━━━━━
Serial No.: {serial_no}
Title: {title}
Node: {next_node}
━━━━━━━━━━━━━━━━━━━━━━━━
"""
    if comment:
        body += f'Approval Note: {comment}\n'
    body += f"""
View Details: http://127.0.0.1:5000/workflow/{workflow_id}

—— DMS Document Management System
"""
    try:
        msg = mime_text.MIMEText(body, 'plain', 'utf-8')
        msg['Subject'] = email.header.Header(subject, 'utf-8')
        msg['From'] = smtp.get('from_addr', 'DMS@dms.local')
        msg['To'] = to_email
        port = int(smtp.get('port', 587))
        if port == 465:
            s = smtplib.SMTP_SSL(smtp['host'], port, timeout=8)
        else:
            s = smtplib.SMTP(smtp['host'], port, timeout=8)
            s.ehlo()
            if smtp.get('use_tls', True):
                s.starttls()
        if smtp.get('username') and smtp.get('password'):
            s.login(smtp['username'], smtp['password'])
        s.sendmail(msg['From'], [to_email], msg.as_string())
        s.quit()
    except Exception as e:
        print(f'[EMAIL ERROR] Failed to send to {to_email}: {e}')


def send_workflow_email(to_email, to_name, serial_no, title, action, next_node, workflow_id, comment='', first_attachment=''):
    """发送ApprovalNotification邮件（异步，Not阻塞主流程）"""
    import threading
    attachment_name = first_attachment or 'No Attachment'
    t = threading.Thread(
        target=_send_email_background,
        args=(to_email, to_name, serial_no, title, action, next_node, workflow_id, comment, attachment_name),
        daemon=True
    )
    t.start()
    return True, '异步发送中'

def create_notification(user_id, notif_type, title, content, workflow_id=None, serial_no=None):
    """CreateShould用内Notification"""
    conn = get_db()
    cur = conn.cursor()
    nid = str(uuid.uuid4())
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    cur.execute('''
        INSERT INTO notifications (id, user_id, type, title, content, workflow_id, serial_no, created_at)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
    ''', (nid, user_id, notif_type, title, content, workflow_id, serial_no, now))
    conn.commit()
    conn.close()

def notify_workflow(workflow_id, action, next_node, approver_id, approver_name, approver_email, comment=''):
    """发送Notification（邮件+Should用内Notification）"""
    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT serial_no, title FROM workflows WHERE id=%s', (workflow_id,))
    row = cur.fetchone()
    
    # GetNo.AttachmentforTitle
    first_attachment = ''
    cur.execute('SELECT original_name FROM workflow_files WHERE workflow_id=%s ORDER BY created_at LIMIT 1', (workflow_id,))
    file_row = cur.fetchone()
    if file_row:
        first_attachment = file_row['original_name']
    
    conn.close()
    if not row:
        return
    serial_no, title = row['serial_no'], row['title']

    # ShouldNotification
    action_content = {
        'Initiated': f'您一个新Approval等待Handle：「{title}」',
        'Approved': f'流程「{serial_no}」Already通过「{next_node}」节点',
        'Returned': f'流程「{serial_no}」Already被Returned至Initiated人',
        '完成': f'流程「{serial_no}」Completed，FilePublished',
    }
    content = action_content.get(action, f'流程「{serial_no}」新动态')
    create_notification(approver_id, 'workflow', title, content, workflow_id, serial_no)

    # Notification（，Not）
    if approver_email:
        try:
            send_workflow_email(approver_email, approver_name, serial_no, title, action, next_node, workflow_id, comment, first_attachment)
        except Exception:
            pass

# =====================================================================
# Data Backup/Restore API
# =====================================================================
import zipfile

# =====================================================================
# Downloads DMSBK Backup File
# =====================================================================
@app.route('/api/backup/list-dmsbk')
@role_required('System Admin', 'Quality Specialist')
def api_list_dmsbk():
    try:
        downloads = os.path.expanduser('~\\Downloads')
        folders = []
        if os.path.isdir(downloads):
            for name in os.listdir(downloads):
                full = os.path.join(downloads, name)
                if name.startswith('DMSBK_') and os.path.isdir(full):
                    folders.append({'name': name, 'path': full})
        folders.sort(key=lambda x: x['name'], reverse=True)
        return jsonify({'success': True, 'folders': folders})
    except Exception as e:
        return jsonify({'success': False, 'folders': [], 'message': str(e)})


# =====================================================================
# File - ReturnPathFileList
# =====================================================================
@app.route('/api/browse-folders')
@role_required('System Admin', 'Quality Specialist')
def api_browse_folders():
    """浏览指定Path下File夹List，for前端File夹Select器"""
    try:
        path = request.args.get('path', '')
        
        # IfPath，ReturnList
        if not path:
            import ctypes
            drives = []
            bitmask = ctypes.windll.kernel32.GetLogicalDrives()
            for letter in range(26):
                if bitmask & (1 << letter):
                    drive = f'{chr(65 + letter)}:\\'
                    if os.path.exists(drive):
                        vol_name = ctypes.create_unicode_buffer(261)
                        result = ctypes.windll.kernel32.GetVolumeInformationW(
                            drive, vol_name, 261, None, None, None, None, 0
                        )
                        vol_label = vol_name.value if result else ''
                        display_name = f'{drive[:2]} {vol_label}' if vol_label else drive[:2]
                        drives.append({
                            'name': display_name,
                            'path': drive,
                            'type': 'drive'
                        })
            return jsonify({'success': True, 'folders': drives, 'current_path': '', 'parent_path': None})
        
        # Path
        path = os.path.normpath(path)
        if not os.path.isdir(path):
            return jsonify({'success': False, 'message': 'Path Not存AtOrNotFile夹'})
        
        # GetDirectory
        parent_path = os.path.dirname(path)
        if parent_path == path:  # Directory
            parent_path = None
        
        # File
        folders = []
        try:
            for name in os.listdir(path):
                full_path = os.path.join(path, name)
                if os.path.isdir(full_path):
                    folders.append({
                        'name': name,
                        'path': full_path,
                        'type': 'folder'
                    })
        except PermissionError:
            pass  # PermissionFile
        
        # Sort：FileAt，NameSort
        folders.sort(key=lambda x: x['name'].lower())
        
        return jsonify({
            'success': True,
            'folders': folders,
            'current_path': path,
            'parent_path': parent_path
        })
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})


# =====================================================================
# FileRefresh
# =====================================================================
def _do_refresh_files():
    """执行File清单Refresh：清理物理Not存At但Data库残留Record（批量Handle）"""
    try:
        # 1. File（，DirectoryOptimize）
        existing = set()
        batch_size = 5000
        if os.path.exists(PUBLISHED_DIR):
            for r, d, fs in os.walk(PUBLISHED_DIR):
                for fn in fs:
                    existing.add(fn)
                # 5000FileHandle，avoid
                if len(existing) > batch_size:
                    break
        
        removed = 0
        conn2 = get_db()
        cur2 = conn2.cursor()
        
        # 2. SearchDelete（Handle1000 ）
        batch = 1000
        while True:
            cur2.execute('SELECT id, filename FROM files WHERE status=%s LIMIT %s', ('Published', batch))
            rows = cur2.fetchall()
            if not rows:
                break
            
            to_delete = [r['id'] for r in rows if r['filename'] not in existing]
            if to_delete:
                placeholders = ','.join(['%s'] * len(to_delete))
                cur2.execute(f'DELETE FROM files WHERE id IN ({placeholders})', to_delete)
                removed += cur2.rowcount
        
        conn2.commit()
        conn2.close()
        return removed
    except Exception as e:
        print(f'[File清单Refresh] Failed: {e}')
        return -1

@app.route('/api/file-list/refresh')
@role_required('System Admin')
def api_refresh_file_list():
    removed = _do_refresh_files()
    if removed >= 0:
        return jsonify({'success': True, 'message': f'File清单AlreadyRefresh，清理 {removed}  残留Record'})
    return jsonify({'success': False, 'message': 'Refresh Failed'})


# =====================================================================
# RestoreSSE
# =====================================================================
@app.route('/api/restore-stream')
def restore_stream():
    """SSE流式推送Restore进度"""
    import time
    # SSE cookie， 'restore' key
    KEY = 'restore'
    def generate():
        last_state = None
        for _ in range(200):  # 80
            state = _restore_progress.get(KEY)
            if state:
                msg = f"data: {state['phase']}|{state['current']}|{state['total']}|{state.get('filename', '')}\n\n"
                yield msg.encode('utf-8')
                if state['phase'] in ('done', 'error'):
                    break
                last_state = state
            else:
                # If done/error Status，
                if last_state and last_state.get('phase') in ('done', 'error'):
                    break
                yield b"data: waiting|0|0|\n\n"
            time.sleep(0.4)
        yield b"data: end|0|0|\n\n"

    resp = Response(generate(), mimetype='text/event-stream')
    resp.headers['Cache-Control'] = 'no-cache'
    resp.headers['X-Accel-Buffering'] = 'no'
    return resp


@app.route('/api/backup', methods=['POST'])
@role_required('System Admin', 'Quality Specialist')
def api_backup():
    data = request.get_json()
    action = data.get('action')

    if action == 'backup':
        import datetime as dt
        now_str = dt.datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # GetUserPath
        target_path = data.get('target_path', '').strip()
        
        if target_path:
            # UserPath，AtPathCreate DMSBK_Time Directory
            if not os.path.isabs(target_path):
                return jsonify({'success': False, 'message': 'Please提供绝对Path，例如 D:\\Backup\\DMS'})
            # ensureDirectoryAt
            os.makedirs(target_path, exist_ok=True)
            bk_dir = os.path.join(target_path, 'DMSBK_' + now_str)
        else:
            # SaveToDownloadDirectory
            download_dir = os.path.expanduser('~\\Downloads')
            bk_dir = os.path.join(download_dir, 'DMSBK_' + now_str)
        
        os.makedirs(bk_dir, exist_ok=True)
        docapr_dir = os.path.join(bk_dir, 'Doc Apr')
        os.makedirs(docapr_dir, exist_ok=True)

        # Backup File（published Directory）
        conn = get_db()
        try:
            cur = conn.cursor()
            # Backup FileFile（Obsolete file），File Name(original_name)Backup
            cur.execute(
                "SELECT id, filename, original_name, department, file_type, file_path "
                "FROM files WHERE is_published=1 AND status != 'Obsolete'"
            )
            files_count = 0
            skipped_missing = 0
            for f in cur.fetchall():
                f = dict_from_row(f)
                dept = f.get('department', '未Category') or '未Category'
                ftype = f.get('file_type', '未Category') or '未Category'
                target_dir = os.path.join(bk_dir, dept, ftype)
                os.makedirs(target_dir, exist_ok=True)
                src = f['file_path']
                if os.path.exists(src):
                    files_count += 1
                    # original_name ForExportFile Name（）
                    orig_name = f.get('original_name') or f['filename']
                    dest = os.path.join(target_dir, orig_name)
                    shutil.copy2(src, dest)
                else:
                    skipped_missing += 1
            cur.execute('SELECT serial_no, title, department, initiator_name, created_at, status FROM workflows ORDER BY created_at DESC')
            workflows_count = 0
            import re as _re
            for w in cur.fetchall():
                w = dict_from_row(w)
                sn = w['serial_no'] or ''
                title = w['title'] or ''
                # Already RestoredRecord（serial_no，OrtitleDoc AprFile Name）
                if '\n' in sn or '\r' in sn:
                    continue
                if sn and title.startswith(sn + '_'):
                    continue
                # Serial No.TitleFile Name
                safe_serial = _re.sub(r'[\\/:*%s"<>|\n\r]', '_', sn)
                safe_title = _re.sub(r'[\\/:*%s"<>|\n\r]', '_', title)
                fname = f"{safe_serial}_{safe_title}.txt"[:200]  # prevent
                fpath = os.path.join(docapr_dir, fname)
                with open(fpath, 'w', encoding='utf-8') as out:
                    out.write(f"Serial No.: {w['serial_no']}\nTitle: {w['title']}\nDepartment: {w['department']}\nInitiated人: {w['initiator_name']}\nDate: {w['created_at']}\nStatus: {w['status']}\n")
                workflows_count += 1
            
            # Backupworkflow PDFsToDoc Apr//Directory
            pdf_count = 0
            if os.path.isdir(WORKFLOW_DIR):
                for serial_folder in os.listdir(WORKFLOW_DIR):
                    src_folder = os.path.join(WORKFLOW_DIR, serial_folder)
                    if not os.path.isdir(src_folder):
                        continue
                    # Directory: Doc Apr//
                    dst_folder = os.path.join(docapr_dir, serial_folder)
                    os.makedirs(dst_folder, exist_ok=True)
                    for pdf_file in os.listdir(src_folder):
                        if pdf_file.endswith('.pdf'):
                            src_pdf = os.path.join(src_folder, pdf_file)
                            dst_pdf = os.path.join(dst_folder, pdf_file)
                            shutil.copy2(src_pdf, dst_pdf)
                            pdf_count += 1
        except Exception as e:
            raise
        finally:
            conn.close()

        # ReturnInformation
        skipped_msg = f'（{skipped_missing}个File磁盘缺失Already跳过）' if skipped_missing else ''
        log_operation(session['user_id'], session['username'], 'Data Backup',
                     f'Backup至: {bk_dir}，{files_count}个File{skipped_msg}，{workflows_count} Record，{pdf_count}个PDF')
        return jsonify({
            'success': True,
            'message': f'Backup completed！\nFile库: {files_count} 个File{skipped_msg}\n approval records: {workflows_count}  \nApprovalPDF: {pdf_count} 个\n\nBackup位置: {bk_dir}',
            'path': bk_dir,
            'files_count': files_count,
            'workflows_count': workflows_count,
            'pdf_count': pdf_count
        })

    elif action == 'restore':
        restore_path = data.get('path', '').strip()
        if not restore_path:
            return jsonify({'success': False, 'message': 'Please select Backup File 夹 Path'})
        
        # IfFileName（NotPath），AtDownloadsDirectory
        if not os.path.isabs(restore_path) or not os.path.isdir(restore_path):
            downloads = os.path.expanduser('~\\Downloads')
            #
            candidate = os.path.join(downloads, restore_path)
            if os.path.isdir(candidate):
                restore_path = candidate
            else:
                #
                if os.path.isdir(downloads):
                    for name in os.listdir(downloads):
                        if restore_path in name and name.startswith('DMSBK_'):
                            candidate = os.path.join(downloads, name)
                            if os.path.isdir(candidate):
                                restore_path = candidate
                                break
        
        if not os.path.isdir(restore_path):
            return jsonify({'success': False, 'message': f'找NotToBackup File夹: {restore_path}\nPleaseensureFile夹AtDownloadsDirectory下，Or手动输入完整Path'})
        
        restored = 0
        updated = 0
        KEY = 'restore'  # And SSE key
        _restore_progress[KEY] = {'phase': 'running', 'current': 0, 'total': 0, 'filename': 'scanning...'}

        # ── ：StatisticsRestoreFileTotal ──
        total_count = 0
        for root, dirs, files in os.walk(restore_path):
            rel = os.path.relpath(root, restore_path)
            if rel.startswith('Doc Apr') or rel == '.' or rel == 'Doc Apr':
                continue
            parts = rel.split(os.sep)
            if len(parts) < 2:
                continue
            total_count += len(files)

        _restore_progress[KEY] = {'phase': 'running', 'current': 0, 'total': total_count, 'filename': ''}

        conn = get_db()
        try:
            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            # ── 1. RestoreFile（ Department+File Type+File Name ，Already At）────────
            processed = 0
            for root, dirs, files in os.walk(restore_path):
                rel = os.path.relpath(root, restore_path)
                if rel.startswith('Doc Apr') or rel == '.' or rel == 'Doc Apr':
                    continue
                parts = rel.split(os.sep)
                if len(parts) < 2:
                    continue
                dept = parts[0]
                ftype = parts[1]
                
                for fname in files:
                    processed += 1
                    _restore_progress[KEY] = {'phase': 'running', 'current': processed, 'total': total_count, 'filename': fname}

                    src = os.path.join(root, fname)
                    ext = os.path.splitext(fname)[1]
                    file_size = os.path.getsize(src)
                    
                    # File NameFor original_name（）
                    real_name = fname
                    
                    # WhetherAlready At Department+Type+original_name Record
                    # 【】Search File（is_published=1 status!='Obsolete'），
                    # ensureRestoreOperationNotObsolete fileRecord
                    cur = conn.cursor()
                    cur.execute('''
                        SELECT id, file_path, is_published, status FROM files
                        WHERE department=%s AND file_type=%s AND original_name=%s
                          AND is_published=1 AND status != 'Obsolete'
                        LIMIT 1
                    ''', (dept, ftype, real_name))
                    existing = cur.fetchone()
                    
                    if existing:
                        # AlreadyRecord → UpdateStatus（PublishedAlready Obsolete，RestoreTopublished）
                        existing_dict = dict_from_row(existing)
                        old_id = existing_dict['id']
                        # RestoreTo PUBLISHED_DIR，NotNeedPath（CanCanobsoleteDirectory）
                        saved_name = f'{uuid.uuid4().hex}{ext}'
                        dept_dir = os.path.join(PUBLISHED_DIR, dept, ftype)
                        os.makedirs(dept_dir, exist_ok=True)
                        dest = os.path.join(dept_dir, saved_name)
                        shutil.copy2(src, dest)
                        file_size = os.path.getsize(dest)
                        # Status，RestoreToPublished
                        cur.execute('''
                            UPDATE files SET original_name=%s, file_path=%s, file_size=%s, status='Published',
                                is_published=1, updated_at=%s, uploader_id=%s, uploader_name=%s
                            WHERE id=%s
                        ''', (real_name, dest, file_size, now, session['user_id'], session['username'], old_id))
                        updated += 1
                    else:
                        # NotAt → 
                        saved_name = f'{uuid.uuid4().hex}{ext}'
                        dept_dir = os.path.join(PUBLISHED_DIR, dept, ftype)
                        os.makedirs(dept_dir, exist_ok=True)
                        dest = os.path.join(dept_dir, saved_name)
                        shutil.copy2(src, dest)
                        file_size = os.path.getsize(dest)
                        cur.execute('''
                            INSERT INTO files (id, filename, original_name, department, file_type,
                                file_path, file_size, status, uploader_id, uploader_name,
                                created_at, updated_at, is_published)
                            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,1)
                        ''', (str(uuid.uuid4()), saved_name, real_name, dept, ftype,
                              dest, file_size, 'Published', session['user_id'],
                              session['username'], now, now))
                        restored += 1
            
            conn.commit()
        except Exception as e:
            conn.rollback()
            raise
        finally:
            conn.close()
        
        total_file = restored + updated
        _restore_progress[KEY] = {"phase": "done", "current": total_file, "total": total_file, "filename": ""}
        log_operation(session['user_id'], session['username'], 'DataRestore',
                     f'From {restore_path} Restore/Update {total_file} 个File（跳过 approval records）')
        return jsonify({
            'success': True,
            'message': f'Restore completed！\n新增File: {restored} 个\n覆盖Update: {updated} 个',
            'restored': restored,
            'updated': updated,
            'total': total_file
        })

    return jsonify({'success': False, 'message': 'Unknown operation'})

# =====================================================================
# DownloadBackup（ZIPTo）
# =====================================================================
@app.route('/api/backup-download')
@role_required('System Admin', 'Quality Specialist')
def api_backup_download():
    """WillBackupData生成ZIPFile并Return给浏览器Download"""
    import datetime as dt
    import io
    import zipfile
    
    try:
        # CreateZIP
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            # 1. Backup File
            conn = get_db()
            cur = conn.cursor()
            cur.execute(
                "SELECT id, filename, original_name, department, file_type, file_path "
                "FROM files WHERE is_published=1 AND status != 'Obsolete'"
            )
            files_count = 0
            for f in cur.fetchall():
                f = dict_from_row(f)

                dept = f.get('department', '未Category') or '未Category'
                ftype = f.get('file_type', '未Category') or '未Category'
                src = f['file_path']
                if os.path.exists(src):
                    files_count += 1
                    orig_name = f.get('original_name') or f['filename']
                    arcname = f'{dept}/{ftype}/{orig_name}'
                    zf.write(src, arcname)
            conn.close()

            # 2. Backup approval recordsForCSV
            conn = get_db()
            cur = conn.cursor()
            cur.execute('SELECT serial_no, title, department, initiator_name, created_at, status FROM workflows ORDER BY created_at DESC')
            import re as _re
            rows = []
            for w in cur.fetchall():
                w = dict_from_row(w)
                sn = w['serial_no'] or ''
                title = w['title'] or ''
                if '\\n' in sn or '\\r' in sn:
                    continue
                if sn and title.startswith(sn + '_'):
                    continue
                rows.append(w)
            conn.close()

            if rows:
                import csv as _csv
                csv_buffer = io.StringIO()
                writer = _csv.DictWriter(csv_buffer, fieldnames=['serial_no', 'title', 'department', 'initiator_name', 'created_at', 'status'])
                writer.writeheader()
                writer.writerows(rows)
                zf.writestr('workflows.csv', csv_buffer.getvalue())

        zip_buffer.seek(0)
        now_str = dt.datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f'DMS_backup_{now_str}.zip'

        return Response(
            zip_buffer.getvalue(),
            mimetype='application/zip',
            headers={
                'Content-Disposition': f'attachment; filename={filename}'
            }
        )
    except Exception as e:
        return jsonify({'success': False, 'message': f'生成BackupFailed: {str(e)}'})

# =====================================================================
# File（FileAndDataRecord）
# =====================================================================
@app.route('/api/files/audit-published', methods=['POST'])
@role_required('System Admin', 'Quality Specialist')
def api_audit_published():
    """发布File库整理：对比 published DirectoryAndData库，双向清理Not一致项"""
    import shutil

    # 1. published DirectoryFile
    physical_files = {}  # filename -> full_path
    for root, dirs, files in os.walk(PUBLISHED_DIR):
        for fname in files:
            physical_files[fname] = os.path.join(root, fname)

    # 2. SearchPublishedRecord
    conn = pymysql.connect(host='localhost', user='dms_user', password='DmsUser123!',
                           database='dms', charset='utf8mb4',
                           cursorclass=pymysql.cursors.DictCursor)
    try:
        cur = conn.cursor()
        cur.execute("SELECT id, filename, original_name, department, file_type, file_path FROM files WHERE is_published=1")
        db_records = cur.fetchall()
    finally:
        conn.close()

    db_filenames = {r['filename'] for r in db_records}
    disk_filenames = set(physical_files.keys())

    # 3. File（Data）
    orphan_files = disk_filenames - db_filenames
    # 4. Record（Data）
    ghost_records = [r for r in db_records if r['filename'] not in disk_filenames]

    deleted_files = []
    deleted_records = []

    # 5. DeleteFile
    for fname in orphan_files:
        fpath = physical_files[fname]
        try:
            os.remove(fpath)
            deleted_files.append({'filename': fname, 'path': fpath})
        except Exception as e:
            deleted_files.append({'filename': fname, 'path': fpath, 'error': str(e)})

    # 6. DeleteRecord（ShouldIndex）
    for rec in ghost_records:
        try:
            conn = pymysql.connect(host='localhost', user='dms_user', password='DmsUser123!',
                                   database='dms', charset='utf8mb4',
                                   cursorclass=pymysql.cursors.DictCursor)
            try:
                cur = conn.cursor()
                # Delete file_index Record
                cur.execute("DELETE FROM file_index WHERE file_id=%s", (rec['filename'],))
                # Delete keyword_index Record
                cur.execute("DELETE FROM keyword_index WHERE file_id=%s", (rec['filename'],))
                # Delete files Record
                cur.execute("DELETE FROM files WHERE id=%s", (rec['id'],))
                conn.commit()
                deleted_records.append({'id': rec['id'], 'filename': rec['filename'],
                                        'original_name': rec['original_name']})
            finally:
                conn.close()
        except Exception as e:
            deleted_records.append({'id': rec['id'], 'filename': rec['filename'],
                                    'original_name': rec['original_name'], 'error': str(e)})

    # 7. Directory
    cleaned_dirs = []
    for root, dirs, files in os.walk(PUBLISHED_DIR, topdown=False):
        if not files and not dirs and root != PUBLISHED_DIR:
            try:
                os.rmdir(root)
                cleaned_dirs.append(root)
            except:
                pass

    # 8. Statistics
    final_disk_count = sum(1 for _, _, fs in os.walk(PUBLISHED_DIR) for _ in fs)
    conn2 = pymysql.connect(host='localhost', user='dms_user', password='DmsUser123!',
                            database='dms', charset='utf8mb4',
                            cursorclass=pymysql.cursors.DictCursor)
    try:
        cur2 = conn2.cursor()
        cur2.execute("SELECT COUNT(*) as cnt FROM files WHERE is_published=1")
        final_db_count = cur2.fetchone()['cnt']
    finally:
        conn2.close()

    return jsonify({
        'success': True,
        'original_disk_count': len(physical_files),
        'original_db_count': len(db_records),
        'orphan_files_deleted': len([f for f in deleted_files if 'error' not in f]),
        'ghost_records_deleted': len([r for r in deleted_records if 'error' not in r]),
        'empty_dirs_cleaned': len(cleaned_dirs),
        'final_disk_count': final_disk_count,
        'final_db_count': final_db_count,
        'details': {
            'deleted_files': deleted_files,
            'deleted_records': deleted_records,
            'cleaned_dirs': cleaned_dirs
        }
    })

# =====================================================================
# FromUploadRestore
# =====================================================================
@app.route('/api/restore-upload', methods=['POST'])
@role_required('System Admin', 'Quality Specialist')
def api_restore_upload():
    """From本地UploadZIPFileRestoreData"""
    import zipfile
    import io
    import csv as _csv
    import uuid
    
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': 'Please select ZIP File'})
    
    file = request.files['file']
    if not file.filename:
        return jsonify({'success': False, 'message': 'File NameCannotFor空'})
    
    if not file.filename.lower().endswith('.zip'):
        return jsonify({'success': False, 'message': 'Please Upload ZIP Format File'})
    
    restored = 0
    updated = 0
    
    try:
        # ZIPFile
        zip_data = file.read()
        zip_buffer = io.BytesIO(zip_data)
        
        with zipfile.ZipFile(zip_buffer, 'r') as zf:
            # 1. RestoreFile
            for name in zf.namelist():
                if name.endswith('/') or name == 'workflows.csv':
                    continue  # DirectoryCSV
                
                # Path: Department/Type/File Name
                parts = name.split('/')
                if len(parts) < 3:
                    continue
                dept = parts[0]
                ftype = parts[1]
                filename = parts[-1]
                
                # FileContent
                content = zf.read(name)
                
                # SaveToDirectory
                target_dir = os.path.join(PUBLISHED_DIR, dept, ftype)
                os.makedirs(target_dir, exist_ok=True)
                target_path = os.path.join(target_dir, filename)
                
                is_new = not os.path.exists(target_path)
                with open(target_path, 'wb') as f:
                    f.write(content)
                
                if is_new:
                    restored += 1
                else:
                    updated += 1
                
                # Update database
                conn = get_db()
                cur = conn.cursor()
                cur.execute('SELECT id FROM files WHERE filename=%s AND department=%s AND file_type=%s', 
                           (filename, dept, ftype))
                existing = cur.fetchone()
                
                if existing:
                    cur.execute('''
                        UPDATE files SET file_path=%s, status='Published', is_published=1, updated_at=NOW() 
                        WHERE filename=%s AND department=%s AND file_type=%s
                    ''', (target_path, filename, dept, ftype))
                else:
                    file_id = str(uuid.uuid4())
                    cur.execute('''
                        INSERT INTO files (id, filename, original_name, file_path, department, file_type, 
                                     status, is_published, created_at, updated_at)
                        VALUES (%s, %s, %s, %s, %s, %s, 'Published', 1, NOW(), NOW())
                    ''', (file_id, filename, filename, target_path, dept, ftype))
                conn.close()

            # 2. Restore approval recordsCSV
            if 'workflows.csv' in zf.namelist():
                csv_data = zf.read('workflows.csv').decode('utf-8')
                csv_reader = _csv.DictReader(io.StringIO(csv_data))
                for row in csv_reader:
                    serial_no = row.get('serial_no', '')
                    title = row.get('title', '')
                    if serial_no and title:
                        conn = get_db()
                        cur = conn.cursor()
                        cur.execute('SELECT id FROM workflows WHERE serial_no=%s', (serial_no,))
                        if not cur.fetchone():
                            wf_id = str(uuid.uuid4())
                            cur.execute('''
                                INSERT INTO workflows (id, serial_no, title, department, file_type, initiator_name, created_at, status)
                                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                            ''', (wf_id, serial_no, title, row.get('department', ''), row.get('file_type', ''), 
                                 row.get('initiator_name', ''), row.get('created_at', ''), row.get('status', '')))
                            conn.commit()
                        conn.close()

        total = restored + updated
        log_operation(session['user_id'], session['username'], 'DataRestore(Upload)',
                     f'From本地UploadRestore/Update {total} 个File')
        return jsonify({
            'success': True,
            'message': f'Restore completed！\n新增File: {restored} 个\n覆盖Update: {updated} 个\n正At后台重建Full-text index...',
            'restored': restored,
            'updated': updated,
            'total': total
        })
        # Full-text index（NotReturn）
        if total > 0:
            import threading, subprocess, sys
            def _async_rebuild():
                try:
                    p = os.path.join(os.path.dirname(__file__), 'system', 'rebuild_index.py')
                    subprocess.run([sys.executable, p, '--auto'], capture_output=True, timeout=7200)
                    print(f'[Restore] Full-text index重建完成')
                except Exception as e:
                    print(f'[Restore] Full-text index重建Failed: {e}')
            threading.Thread(target=_async_rebuild, daemon=True).start()
    except Exception as e:
        return jsonify({'success': False, 'message': f'Restore Failed: {str(e)}'})

# =====================================================================
# Department/File Typepage
# =====================================================================
@app.route('/dept-type-manage')
@role_required('System Admin')
def dept_type_manage():
    sidebar = get_sidebar_data(session['role'], session['user_id'])
    sidebar['user_theme'] = session.get('user_theme', 'tech-blue')
    return render_template('dept_type_manage.html', sidebar=sidebar, page_title='Department And File Type Management')

# =====================================================================
# Operation Logpage
# =====================================================================
@app.route('/logs')
@role_required('System Admin', 'Quality Specialist')
def operation_logs():
    page = _int_param(request.args.get('page'), 1)
    per_page = 50
    offset = (page - 1) * per_page

    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT COUNT(*) AS cnt FROM operation_logs')
    total = cur.fetchone()['cnt']
    
    # Pagination
    total_pages = (total + per_page - 1) // per_page if total > 0 else 1
    
    # page（Whenpage2page）
    page_start = max(1, page - 2)
    page_end = min(total_pages, page + 2)
    page_range = list(range(page_start, page_end + 1))
    
    cur.execute('''
        SELECT * FROM operation_logs
        ORDER BY created_at DESC LIMIT %s OFFSET %s
    ''', (per_page, offset))
    logs = [dict_from_row(r) for r in cur.fetchall()]
    conn.close()

    sidebar = get_sidebar_data(session['role'], session['user_id'])
    sidebar['user_theme'] = session.get('user_theme', 'tech-blue')
    return render_template('logs.html',
        logs=logs, total=total, page=page, total_pages=total_pages, page_range=page_range,
        sidebar=sidebar, page_title='Operation Log')

# =====================================================================
# Operation Log API
# =====================================================================
@app.route('/api/logs/clear', methods=['POST'])
@role_required('System Admin')
def api_clear_logs():
    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT COUNT(*) AS cnt FROM operation_logs')
    count = cur.fetchone()['cnt']
    cur.execute('DELETE FROM operation_logs')
    conn.commit()
    conn.close()
    
    log_operation(session['user_id'], session['username'], '清除Log', f'清除 {count}  Operation Log')
    return jsonify({'success': True, 'deleted_count': count})

# =====================================================================
# Configpage
# =====================================================================
@app.route('/system-config')
@role_required('System Admin')
def system_config():
    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT content, created_at FROM announcements ORDER BY created_at DESC LIMIT 1')
    ann = cur.fetchone()
    announcement = dict_from_row(ann) if ann else None
    cur.execute('SELECT `value` FROM config WHERE `key`=%s', ('system_version',))
    ver = cur.fetchone()
    system_version = ver['value'] if ver else 'V1.0.0'
    conn.close()

    sidebar = get_sidebar_data(session['role'], session['user_id'])
    sidebar['user_theme'] = session.get('user_theme', 'tech-blue')
    return render_template('system_config.html',
        announcement=announcement, system_version=system_version,
        sidebar=sidebar, page_title='System Config')

# =====================================================================
# API
# =====================================================================
@app.route('/api/announcement', methods=['GET', 'POST'])
@role_required('System Admin')
def api_announcement():
    conn = get_db()
    cur = conn.cursor()
    if request.method == 'GET':
        cur.execute('SELECT content, created_at FROM announcements ORDER BY created_at DESC LIMIT 1')
        ann = cur.fetchone()
        conn.close()
        return jsonify(dict_from_row(ann) if ann else {})

    data = request.get_json()
    content = data.get('content', '').strip()
    if not content:
        conn.close()
        return jsonify({'success': False, 'message': '公告ContentCannotFor空'})
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    cur.execute('SELECT id FROM announcements ORDER BY created_at DESC LIMIT 1')
    existing = cur.fetchone()
    if existing:
        cur.execute('UPDATE announcements SET content=%s, created_at=%s WHERE id=%s',
                    (content, now, existing['id']))
    else:
        ann_id = str(uuid.uuid4())
        cur.execute('INSERT INTO announcements (id, content, author_id, created_at) VALUES (%s,%s,%s,%s)',
                    (ann_id, content, session['user_id'], now))
    conn.commit()
    conn.close()
    log_operation(session['user_id'], session['username'], 'Update公告', f'公告: {content[:30]}')
    return jsonify({'success': True, 'message': '公告AlreadySave'})

# =====================================================================
# SMTP Config API
# =====================================================================
@app.route('/api/smtp', methods=['GET', 'POST'])
@role_required('System Admin')
def api_smtp():
    if request.method == 'GET':
        smtp = get_smtp_config()
        # NotReturnPassword
        smtp.pop('password', None)
        return jsonify(smtp)

    data = request.get_json()
    smtp = {
        'enabled': bool(data.get('enabled')),
        'host': data.get('host', ''),
        'port': str(data.get('port', 587)),
        'username': data.get('username', ''),
        'password': data.get('password', ''),
        'from_addr': data.get('from_addr', ''),
        'use_tls': bool(data.get('use_tls', True)),
    }
    try:
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            cfg = _json.load(f)
    except Exception:
        cfg = {}
    cfg['smtp'] = smtp
    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        _json.dump(cfg, f, ensure_ascii=False, indent=2)
    log_operation(session['user_id'], session['username'], 'ConfigSMTP', f'SMTP: {smtp["host"]}')
    return jsonify({'success': True, 'message': 'SMTPConfigAlreadySave'})


@app.route('/api/smtp/test', methods=['POST'])
@role_required('System Admin', 'Quality Specialist')
def api_smtp_test():
    """发送Test邮件To指定Email"""
    data = request.get_json() or {}
    test_email = (data.get('test_email') or '').strip()
    if not test_email or '@' not in test_email:
        return jsonify({'success': False, 'message': 'Please enter 效Test 收件 Email Address'}), 400

    smtp = get_smtp_config()
    if not smtp.get('enabled') or not smtp.get('host'):
        return jsonify({'success': False, 'message': 'SMTP is not enabled or not configured. Please save the configuration first.'}), 400

    subject = '[DMS Test Email] Email Notification Feature Verification'
    body = f"""Hello, this is a test email from the DMS Document Management System.

If you can read this email, it means the email notification feature is configured correctly.

━━━━━━━━━━━━━━━━━━━━━━━━
SMTP Server: {smtp.get('host')}:{smtp.get('port')}
Sender Address: {smtp.get('from_addr') or 'Not set'}
Send Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
━━━━━━━━━━━━━━━━━━━━━━━━

—— DMS Document Management System
"""
    try:
        msg = mime_text.MIMEText(body, 'plain', 'utf-8')
        msg['Subject'] = email.header.Header(subject, 'utf-8')
        msg['From'] = smtp.get('from_addr', 'DMS@dms.local')
        msg['To'] = test_email
        port = int(smtp.get('port', 587))
        if port == 465:
            s = smtplib.SMTP_SSL(smtp['host'], port, timeout=10)
        else:
            s = smtplib.SMTP(smtp['host'], port, timeout=10)
            s.ehlo()
            if smtp.get('use_tls', True):
                s.starttls()
        if smtp.get('username') and smtp.get('password'):
            s.login(smtp['username'], smtp['password'])
        s.sendmail(msg['From'], [test_email], msg.as_string())
        s.quit()
        log_operation(session['user_id'], session['username'], '发送Test邮件', f'收件: {test_email}')
        return jsonify({'success': True, 'message': f'Test邮件AlreadySuccess发送至 {test_email}，Please查收。'})
    except smtplib.SMTPAuthenticationError:
        return jsonify({'success': False, 'message': '认证Failed：UsernameOrPassword/授权码Error'}), 400
    except smtplib.SMTPRecipientsRefused:
        return jsonify({'success': False, 'message': '收件人被服务器拒绝，PleaseConfirm收件EmailAddress效'}), 400
    except smtplib.SMTPSenderRefused:
        return jsonify({'success': False, 'message': '发件Address被服务器拒绝，PleaseAtConfig中填写正确发件Address'}), 400
    except smtplib.SMTPException as e:
        return jsonify({'success': False, 'message': f'SMTPError：{str(e)}'}), 400
    except Exception as e:
        return jsonify({'success': False, 'message': f'发送Failed：{str(e)}'}), 500


# =====================================================================
# Config API
# =====================================================================
@app.route('/api/scheduler-config', methods=['GET', 'POST'])
@role_required('System Admin')
def api_scheduler_config():
    """Get/Settings定时任务开关Config"""
    if request.method == 'GET':
        cfg = get_scheduler_config()
        return jsonify(cfg)

    data = request.get_json() or {}
    enabled = bool(data.get('weekly_index_enabled', False))
    cleanup_enabled = bool(data.get('weekly_cleanup_enabled', False))
    try:
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            cfg = _json.load(f)
    except Exception:
        cfg = {}
    if 'scheduler' not in cfg:
        cfg['scheduler'] = {}
    cfg['scheduler']['weekly_index_enabled'] = enabled
    cfg['scheduler']['weekly_cleanup_enabled'] = cleanup_enabled
    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        _json.dump(cfg, f, ensure_ascii=False, indent=2)
    action = '开启' if enabled else '关闭'
    cleanup_action = '开启' if cleanup_enabled else '关闭'
    log_operation(session['user_id'], session['username'], 'Config定时任务',
                  f'每周日IndexUpdate：{action}，Preview清理：{cleanup_action}')
    return jsonify({'success': True, 'message': f'定时任务ConfigAlreadySave（IndexUpdate: {action}, Preview清理: {cleanup_action}）'})


# =====================================================================
# Notification
# =====================================================================
@app.route('/notifications')
@login_required
def notifications():
    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
        SELECT id, type, title, content, workflow_id, serial_no, is_read, created_at
        FROM notifications WHERE user_id=%s ORDER BY created_at DESC LIMIT 100
    ''', (session['user_id'],))
    notifs = [dict_from_row(r) for r in cur.fetchall()]
    # Unread
    cur.execute('SELECT COUNT(*) AS cnt FROM notifications WHERE user_id=%s AND is_read=0', (session['user_id'],))
    unread = cur.fetchone()['cnt']
    conn.close()

    sidebar = get_sidebar_data(session['role'], session['user_id'])
    sidebar['user_theme'] = session.get('user_theme', 'tech-blue')
    sidebar['unread_notifs'] = unread
    return render_template('notifications.html',
        notifications=notifs, unread=unread,
        sidebar=sidebar, page_title='Notification Center')

@app.route('/api/notifications/mark-read', methods=['POST'])
@login_required
def api_mark_read():
    data = request.get_json()
    nid = data.get('id')
    conn = get_db()
    cur = conn.cursor()
    if nid:
        cur.execute('UPDATE notifications SET is_read=1 WHERE id=%s AND user_id=%s', (nid, session['user_id']))
    else:
        cur.execute('UPDATE notifications SET is_read=1 WHERE user_id=%s', (session['user_id'],))
    conn.commit()
    conn.close()
    return jsonify({'success': True})

# =====================================================================
# Logo Upload
# =====================================================================
# =====================================================================
# （Login）
# =====================================================================
@app.route('/public/logo')
def public_logo():
    """公开LogoSearch接口，供Loginpage面使用"""
    logo_path = os.path.join(SYSTEM_DIR, 'static', 'logo.png')
    exists = os.path.exists(logo_path)
    return jsonify({'exists': exists})

# =====================================================================
# Logo（NeedAdminPermission）
# =====================================================================
@app.route('/api/logo', methods=['GET', 'POST', 'DELETE'])
@role_required('System Admin')
def api_logo():
    logo_path = os.path.join(SYSTEM_DIR, 'static', 'logo.png')

    if request.method == 'GET':
        exists = os.path.exists(logo_path)
        return jsonify({'exists': exists})

    elif request.method == 'POST':
        f = request.files.get('logo')
        if not f:
            return jsonify({'success': False, 'message': 'Please select Photo File'})
        allowed_exts = {'.png', '.jpg', '.jpeg', '.gif', '.webp', '.svg'}
        ext = os.path.splitext(f.filename)[1].lower()
        if ext not in allowed_exts:
            return jsonify({'success': False, 'message': '仅Supports PNG/JPG/GIF/WebP/SVG Format'})
        os.makedirs(os.path.dirname(logo_path), exist_ok=True)
        f.save(logo_path)
        log_operation(session['user_id'], session['username'], 'UploadLogo', '更换系统Logo')
        return jsonify({'success': True, 'message': 'Logo AlreadyUpload'})

    elif request.method == 'DELETE':
        if os.path.exists(logo_path):
            os.remove(logo_path)
            log_operation(session['user_id'], session['username'], 'DeleteLogo', 'Delete系统Logo')
        return jsonify({'success': True, 'message': 'Logo Already Deleted'})

# =====================================================================
# LoginpageTitleSettings
# =====================================================================
LOGIN_TITLE_FILE = os.path.join(SYSTEM_DIR, 'login_title.txt')

@app.route('/api/login-title', methods=['GET', 'POST'])
@role_required('System Admin')
def api_login_title():
    if request.method == 'GET':
        title = ''
        if os.path.exists(LOGIN_TITLE_FILE):
            with open(LOGIN_TITLE_FILE, 'r', encoding='utf-8') as f:
                title = f.read().strip()
        return jsonify({'title': title})
    
    elif request.method == 'POST':
        data = request.get_json()
        title = data.get('title', '').strip()
        with open(LOGIN_TITLE_FILE, 'w', encoding='utf-8') as f:
            f.write(title)
        log_operation(session['user_id'], session['username'], '修改LoginTitle', f'Settings For: {title}')
        return jsonify({'success': True, 'message': 'Title Already Saved'})

@app.route('/public/login-title')
def public_login_title():
    """公开LoginTitleSearch接口，供Loginpage面使用"""
    title = '科技公司File Management系统'
    if os.path.exists(LOGIN_TITLE_FILE):
        with open(LOGIN_TITLE_FILE, 'r', encoding='utf-8') as f:
            saved_title = f.read().strip()
            if saved_title:
                title = saved_title
    return jsonify({'title': title})

# =====================================================================
# Data Backup
# =====================================================================
@app.route('/backup')
@role_required('System Admin', 'Quality Specialist')
def backup_page():
    sidebar = get_sidebar_data(session['role'], session['user_id'])
    sidebar['user_theme'] = session.get('user_theme', 'tech-blue')
    return render_template('backup.html', sidebar=sidebar, page_title='Data Backup And Restore')

# =====================================================================
# Department/File Type API
# =====================================================================

@app.route('/skin-settings')
@login_required
def skin_settings_page():
    """Theme Settingspage面"""
    sidebar = get_sidebar_data(session['role'], session['user_id'])
    sidebar['user_theme'] = session.get('user_theme', 'tech-blue')
    return render_template('skin_settings.html', sidebar=sidebar, page_title='Theme Settings')


@app.route('/api/user/theme', methods=['POST'])
@login_required
def save_user_theme():
    """SaveUser主题SettingsToData库"""
    data = request.get_json()
    theme = data.get('theme', 'tech-blue')
    
    conn = get_db()
    cur = conn.cursor()
    cur.execute('UPDATE users SET user_theme=%s WHERE id=%s', (theme, session['user_id']))
    conn.commit()
    conn.close()
    
    # Also update session
    session['user_theme'] = theme
    
    return jsonify({'success': True, 'theme': theme})

@app.route('/api/departments', methods=['GET', 'POST', 'PUT', 'DELETE'])
@login_required
def api_departments():
    conn = get_db()
    cur = conn.cursor()
    method = request.method

    if method == 'GET':
        # GETPleaseLoginUser
        cur.execute('SELECT id, name, sort_order FROM departments ORDER BY sort_order')
        return jsonify([dict_from_row(r) for r in cur.fetchall()])
    
    # OperationSystem Admin
    if session.get('role') != 'System Admin':
        conn.close()
        return jsonify({'success': False, 'message': 'Permission Not足，需NeedSystem AdminPermission'}), 403

    elif method == 'POST':
        data = request.get_json()
        if not data.get('name'):
            conn.close()
            return jsonify({'success': False, 'message': 'Department Name Cannot Be Empty空'})
        cur.execute('SELECT id FROM departments WHERE name=%s', (data['name'],))
        if cur.fetchone():
            conn.close()
            return jsonify({'success': False, 'message': 'Department Already存At'})
        cur.execute('SELECT MAX(sort_order) AS max_val FROM departments')
        max_order = int(cur.fetchone()['max_val'] or 0)
        dept_id = str(uuid.uuid4())
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        cur.execute('''
            INSERT INTO departments (id, name, sort_order, created_at)
            VALUES (%s,%s,%s,%s)
        ''', (dept_id, data['name'].strip(), max_order + 1, now))
        # ForDepartmentInitialize 6 File Type
        cur.execute("SELECT id, name FROM file_types ORDER BY sort_order")
        for i, ft in enumerate(cur.fetchall(), 1):
            cur.execute(
                "INSERT IGNORE INTO dept_file_types (id, department, name, sort_order, created_at) VALUES (%s,%s,%s,%s,%s)",
                (str(uuid.uuid4()), data['name'].strip(), ft['name'], i, now)
            )
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'Department Already添加'})

    elif method == 'PUT':
        data = request.get_json()
        dept_id = data.get('id')
        action = data.get('action', '')
        if action == 'move':
            direction = data.get('direction')
            if not dept_id or direction not in ('up', 'down'):
                conn.close()
                return jsonify({'success': False, 'message': '参数Error'})
            cur.execute('SELECT sort_order FROM departments WHERE id=%s', (dept_id,))
            row = cur.fetchone()
            if not row:
                conn.close()
                return jsonify({'success': False, 'message': 'Record Not存At'})
            current_order = row['sort_order']
            if direction == 'up':
                cur.execute(
                    'SELECT id, sort_order FROM departments WHERE sort_order < %s ORDER BY sort_order DESC LIMIT 1',
                    (current_order,)
                )
            else:
                cur.execute(
                    'SELECT id, sort_order FROM departments WHERE sort_order > %s ORDER BY sort_order ASC LIMIT 1',
                    (current_order,)
                )
            neighbor = cur.fetchone()
            if not neighbor:
                conn.close()
                return jsonify({'success': True, 'message': 'Already At最顶端' if direction == 'up' else 'Already At最底端'})
            cur.execute('UPDATE departments SET sort_order=%s WHERE id=%s', (neighbor['sort_order'], dept_id))
            cur.execute('UPDATE departments SET sort_order=%s WHERE id=%s', (current_order, neighbor['id']))
            conn.commit()
            conn.close()
            return jsonify({'success': True, 'message': 'Already移动'})
        else:
            name = data.get('name', '').strip()
            order = data.get('sort_order', 0)
            if not dept_id or not name:
                conn.close()
                return jsonify({'success': False, 'message': '参数Error'})
            cur.execute('UPDATE departments SET name=%s, sort_order=%s WHERE id=%s',
                       (name, order, dept_id))
            conn.commit()
            conn.close()
            return jsonify({'success': True, 'message': 'Department Already Updated'})

    elif method == 'DELETE':
        data = request.get_json()
        dept_id = data.get('id')
        if not dept_id:
            conn.close()
            return jsonify({'success': False, 'message': '缺少DepartmentID'})
        cur.execute('DELETE FROM departments WHERE id=%s', (dept_id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'Department Already Deleted'})

# File TypeSort
FILE_TYPE_ORDER = ['Manual', 'Procedure', 'Work Instruction', 'Form', 'External Documents', 'Record', 'Others']

def _resort_file_types(cur):
    """按固定顺序重排所File Typesort_order"""
    cur.execute('SELECT id, name FROM file_types ORDER BY sort_order')
    rows = cur.fetchall()
    for row in rows:
        name = row['name'] if hasattr(row, 'keys') else row[1]
        if name in FILE_TYPE_ORDER:
            order = FILE_TYPE_ORDER.index(name) + 1
        else:
            # TypeAt，
            order = len(FILE_TYPE_ORDER) + 1 + (list(rows).index(row))
        fid = row['id'] if hasattr(row, 'keys') else row[0]
        cur.execute('UPDATE file_types SET sort_order=%s WHERE id=%s', (order, fid))
    # dept_file_types
    cur.execute('SELECT DISTINCT department FROM dept_file_types')
    depts = [r['department'] for r in cur.fetchall()]
    for dept in depts:
        cur.execute('SELECT id, name FROM dept_file_types WHERE department=%s', (dept,))
        for row in cur.fetchall():
            name = row['name'] if hasattr(row, 'keys') else row[1]
            if name in FILE_TYPE_ORDER:
                order = FILE_TYPE_ORDER.index(name) + 1
            else:
                order = 999
            fid = row['id'] if hasattr(row, 'keys') else row[0]
            cur.execute('UPDATE dept_file_types SET sort_order=%s WHERE id=%s', (order, fid))

@app.route('/api/file-types', methods=['GET', 'POST', 'PUT', 'DELETE'])
@login_required
def api_file_types():
    conn = get_db()
    cur = conn.cursor()
    method = request.method

    if method == 'GET':
        # GETPleaseLoginUser
        cur.execute('SELECT id, name, sort_order FROM file_types ORDER BY sort_order')
        return jsonify([dict_from_row(r) for r in cur.fetchall()])
    
    # OperationSystem Admin
    if session.get('role') != 'System Admin':
        conn.close()
        return jsonify({'success': False, 'message': 'Permission Not足，需NeedSystem AdminPermission'}), 403

    elif method == 'POST':
        data = request.get_json()
        if not data.get('name'):
            conn.close()
            return jsonify({'success': False, 'message': 'Type Name Cannot Be Empty空'})
        name = data['name'].strip()
        cur.execute('SELECT id FROM file_types WHERE name=%s', (name,))
        if cur.fetchone():
            conn.close()
            return jsonify({'success': False, 'message': 'Type Already存At'})
        ft_id = str(uuid.uuid4())
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        cur.execute('INSERT INTO file_types (id, name, sort_order, created_at) VALUES (%s,%s,%s,%s)',
                   (ft_id, name, 999, now))
        # Re-sort all by fixed order
        _resort_file_types(cur)
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'File Type Already添加'})

    elif method == 'PUT':
        data = request.get_json()
        ft_id = data.get('id')
        name = data.get('name', '').strip()
        if not ft_id or not name:
            conn.close()
            return jsonify({'success': False, 'message': '参数Error'})
        cur.execute('UPDATE file_types SET name=%s WHERE id=%s', (name, ft_id))
        # Re-sort all by fixed order
        _resort_file_types(cur)
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'File Type Already Updated'})

    elif method == 'DELETE':
        data = request.get_json()
        ft_id = data.get('id')
        if not ft_id:
            conn.close()
            return jsonify({'success': False, 'message': '缺少ID'})
        cur.execute('DELETE FROM file_types WHERE id=%s', (ft_id,))
        # Re-sort all by fixed order
        _resort_file_types(cur)
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'File Type Already Deleted'})

@app.route('/api/dept-file-types', methods=['GET', 'POST', 'PUT', 'DELETE'])
@role_required('System Admin')
def api_dept_file_types():
    """按DepartmentFile Type管理：Supports每个Department独立ConfigType和Sort"""
    conn = get_db()
    cur = conn.cursor()
    method = request.method

    if method == 'GET':
        dept = request.args.get('dept', '')
        # GetCanType（ 6 ）
        cur.execute('SELECT name FROM file_types ORDER BY sort_order')
        available = [r['name'] for r in cur.fetchall()]
        if dept:
            cur.execute(
                'SELECT id, department, name, sort_order FROM dept_file_types WHERE department=%s ORDER BY sort_order',
                (dept,)
            )
            rows = [dict_from_row(r) for r in cur.fetchall()]
            # Department AlreadyType
            owned = {r['name'] for r in rows}
        else:
            rows = []
            owned = set()
        conn.close()
        return jsonify({
            'owned': rows,
            'available': [t for t in available if t not in owned]
        })

    elif method == 'POST':
        data = request.get_json()
        dept = data.get('department', '').strip()
        name = data.get('name', '').strip()
        if not dept or not name:
            conn.close()
            return jsonify({'success': False, 'message': 'Department和Type Name Cannot Be Empty空'})
        # file_types 
        cur.execute('SELECT id FROM file_types WHERE name=%s', (name,))
        if not cur.fetchone():
            conn.close()
            return jsonify({'success': False, 'message': '无效File Type'})
        # Already At
        cur.execute(
            'SELECT id FROM dept_file_types WHERE department=%s AND name=%s',
            (dept, name)
        )
        if cur.fetchone():
            conn.close()
            return jsonify({'success': False, 'message': '该Department Already拥此Type'})
        cur.execute('SELECT MAX(sort_order) AS max_val FROM dept_file_types WHERE department=%s', (dept,))
        max_order = int(cur.fetchone()['max_val'] or 0)
        ft_id = str(uuid.uuid4())
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        cur.execute(
            'INSERT INTO dept_file_types (id, department, name, sort_order, created_at) VALUES (%s,%s,%s,%s,%s)',
            (ft_id, dept, name, 999, now)
        )
        # Re-sort by fixed order
        _resort_file_types(cur)
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': f'Already For「{dept}」添加「{name}」'})

    elif method == 'PUT':
        data = request.get_json()
        ft_id = data.get('id')
        # Edit（Name）- Sort
        name = data.get('name', '').strip()
        if not ft_id or not name:
            conn.close()
            return jsonify({'success': False, 'message': '参数Error'})
        cur.execute('UPDATE dept_file_types SET name=%s WHERE id=%s', (name, ft_id))
        # Re-sort by fixed order
        _resort_file_types(cur)
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'Already Updated'})

    elif method == 'DELETE':
        data = request.get_json()
        ft_id = data.get('id')
        if not ft_id:
            conn.close()
            return jsonify({'success': False, 'message': '缺少ID'})
        cur.execute('DELETE FROM dept_file_types WHERE id=%s', (ft_id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'Already Deleted'})

# =====================================================================
#
# =====================================================================
@app.route('/api/theme', methods=['POST'])
@login_required
def api_theme():
    data = request.get_json()
    theme = data.get('theme', 'tech-blue')
    conn = get_db()
    cur = conn.cursor()
    cur.execute('UPDATE users SET user_theme=%s WHERE id=%s',
                 (theme, session['user_id']))
    conn.commit()
    conn.close()
    session['user_theme'] = theme
    return jsonify({'success': True})

# =====================================================================
# Information
# =====================================================================
@app.route('/api/system-info')
@login_required
def api_system_info():
    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT COUNT(*) AS cnt FROM files')
    total_files = cur.fetchone()['cnt']
    cur.execute('SELECT COUNT(*) AS cnt FROM workflows')
    total_workflows = cur.fetchone()['cnt']
    cur.execute('SELECT COUNT(*) AS cnt FROM users WHERE status="Active"')
    active_users = cur.fetchone()['cnt']
    conn.close()
    return jsonify({
        'version': 'V1.0.0',
        'total_files': total_files,
        'total_workflows': total_workflows,
        'active_users': active_users
    })

# =====================================================================
# Check（Login，）
# =====================================================================
@app.route('/api/health')
def api_health():
    """健康Check端点：检测Data库连接和磁盘空间"""
    checks = {'status': 'ok', 'timestamp': datetime.now().isoformat()}
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute('SELECT 1 AS test')
        cur.fetchone()
        conn.close()
        checks['database'] = 'ok'
    except Exception as e:
        checks['database'] = f'error: {str(e)[:100]}'
        checks['status'] = 'degraded'
    try:
        disk_usage = shutil.disk_usage(BASE_DIR)
        disk_free_gb = round(disk_usage.free / (1024**3), 2)
        disk_total_gb = round(disk_usage.total / (1024**3), 2)
        checks['disk'] = {'free_gb': disk_free_gb, 'total_gb': disk_total_gb, 'usage_pct': round((1 - disk_usage.free / disk_usage.total) * 100, 1)}
        if disk_usage.free < 1024**3:  # < 1GB
            checks['status'] = 'degraded'
    except Exception:
        checks['disk'] = 'unknown'
    code = 200 if checks['status'] == 'ok' else 503
    return jsonify(checks), code

# =====================================================================
# FileSearch
# =====================================================================
def _files_access_sql(role):
    """ReturnFile访问控制SQL 件片段和参数"""
    if role not in ["System Admin", "Quality Specialist"]:
        return "is_published=1 AND status=%s", ["Published"]
    return "status!='Obsolete'", []


@app.route('/api/files/search')
@login_required
def api_files_search():
    q = request.args.get("q", "").strip()
    dept = request.args.get("dept", "").strip()
    cat = request.args.get("cat", "").strip()
    fulltext = request.args.get("fulltext", "0").strip()
    if not q:
        return jsonify({'results': [], 'count': 0})

    role = session.get("role", "")
    conn = get_db()
    cur = conn.cursor()

    p = "%" + q + "%"
    where_parts = []
    params = []

    where_parts.append("(original_name LIKE %s OR department LIKE %s OR file_type LIKE %s OR filename LIKE %s)")
    params.extend([p, p, p, p])

    if dept:
        where_parts.append("department = %s")
        params.append(dept)

    if cat:
        where_parts.append("file_type = %s")
        params.append(cat)

    access_sql, access_params = _files_access_sql(role)
    where_parts.append(access_sql)
    params.extend(access_params)

    where_sql = " AND ".join(where_parts)
    
    # SearchTotal
    count_query = "SELECT COUNT(*) as cnt FROM files WHERE " + where_sql
    cur.execute(count_query, params)
    total_count = cur.fetchone()['cnt']
    
    # SearchResult（100 ，avoidData）
    query = "SELECT id, filename, original_name, department, file_type, file_size, uploader_name, created_at, status FROM files WHERE " + where_sql + " ORDER BY created_at DESC LIMIT 100"
    cur.execute(query, params)
    rows = cur.fetchall()

    results = []
    seen_ids = set()

    for r in rows:
        f = dict_from_row(r)
        seen_ids.add(f['id'])
        size = f.get('file_size', 0) or 0
        if size < 1024:
            f['size_fmt'] = str(size) + "B"
        elif size < 1024*1024:
            f['size_fmt'] = str(round(size/1024, 1)) + "KB"
        else:
            f['size_fmt'] = str(round(size/1024/1024, 1)) + "MB"
        f['icon'] = file_ext_to_icon(f['original_name'])
        hl = "<mark style=\"background:rgba(0,212,255,0.3);color:inherit;border-radius:2px;padding:0 2px;\">" + q + "</mark>"
        f['highlight_name'] = f['original_name'].replace(q, hl)
        f['match_type'] = 'name'
        results.append(f)

    # Search：Search File Content
    if fulltext == '1':
        ft_access_sql, ft_access_params = _files_access_sql(role)
        # 1: keyword_index Index（，Not ft_min_word_len ）
        kw_rows = []
        try:
            kw_query = ("SELECT ki.file_id, fi.content, f.id, f.original_name, f.department, f.file_type, "
                        "f.file_size, f.uploader_name, f.created_at, f.status, f.filename "
                        "FROM keyword_index ki "
                        "JOIN file_index fi ON ki.file_id = fi.file_id "
                        "LEFT JOIN files f ON ki.file_id = f.filename "
                        "WHERE ki.term = %s AND " + ft_access_sql + " LIMIT 100")
            cur.execute(kw_query, [q] + ft_access_params)
            kw_rows = cur.fetchall()
        except Exception as e:
            print(f'[Search] keyword_indexSearch Failed: {e}')

        # 2: FULLTEXT Index（1Result）
        ft_rows = []
        exclude_ids = [r['file_id'] for r in kw_rows]
        try:
            exclude_clause = ''
            exclude_params = []
            if exclude_ids:
                placeholders = ','.join(['%s'] * len(exclude_ids))
                exclude_clause = f"AND fi.file_id NOT IN ({placeholders})"
                exclude_params = exclude_ids
            ft_query = ("SELECT fi.file_id, fi.content, f.id, f.original_name, f.department, f.file_type, "
                        "f.file_size, f.uploader_name, f.created_at, f.status, f.filename "
                        "FROM file_index fi JOIN files f ON fi.file_id=f.filename "
                        "WHERE MATCH(fi.content) AGAINST(%s IN BOOLEAN MODE) AND " + ft_access_sql + " " + exclude_clause + " LIMIT 100")
            cur.execute(ft_query, [q] + ft_access_params + exclude_params)
            ft_rows = cur.fetchall()
        except Exception as e:
            print(f'[Search] FULLTEXTSearch Failed，回退LIKE: {e}')
            import re as _re_ft
            if _re_ft.match(r'^[\u4e00-\u9fa5]+$', q):
                chars = list(q)
                like_p = '%'.join(chars)
            else:
                like_p = q
            exclude_clause2 = ''
            exclude_params2 = []
            if exclude_ids:
                placeholders2 = ','.join(['%s'] * len(exclude_ids))
                exclude_clause2 = f"AND fi.file_id NOT IN ({placeholders2})"
                exclude_params2 = exclude_ids
            ft_query = ("SELECT fi.file_id, fi.content, f.id, f.original_name, f.department, f.file_type, "
                        "f.file_size, f.uploader_name, f.created_at, f.status, f.filename "
                        "FROM file_index fi JOIN files f ON fi.file_id=f.filename "
                        "WHERE fi.content LIKE %s AND " + ft_access_sql + " " + exclude_clause2 + " LIMIT 100")
            cur.execute(ft_query, ["%" + like_p + "%"] + ft_access_params + exclude_params2)
            ft_rows = cur.fetchall()
        # 12Result
        all_ft_rows = list(kw_rows)
        for r in ft_rows:
            if r['file_id'] not in exclude_ids:
                all_ft_rows.append(r)
                exclude_ids.append(r['file_id'])
        for r in all_ft_rows:
            fid = r['file_id']
            if fid in seen_ids:
                continue
            seen_ids.add(fid)
            f = dict(r)
            f['id'] = r['id']  # UUID（f.id），NotNeedFile Name
            f['icon'] = file_ext_to_icon(f['original_name'])
            hl = "<mark style=\"background:rgba(255,200,0,0.3);color:inherit;border-radius:2px;padding:0 2px;\">" + q + "</mark>"
            f['highlight_name'] = f['original_name'].replace(q, hl)
            f['match_type'] = 'content'
            content_snippet = (r['content'] or '')[:300]
            f['content_snippet'] = content_snippet.replace(q, hl)
            size = f.get('file_size', 0) or 0
            if size < 1024:
                f['size_fmt'] = str(size) + "B"
            elif size < 1024*1024:
                f['size_fmt'] = str(round(size/1024, 1)) + "KB"
            else:
                f['size_fmt'] = str(round(size/1024/1024, 1)) + "MB"
            results.append(f)

    conn.close()
    return jsonify({'results': results, 'count': len(results), 'total_count': total_count, 'limited': total_count > 200 or len(results) >= 100})


# =====================================================================
#
# =====================================================================
if __name__ == '__main__':
    print('='*60)
    print('  DMS File Management系统 (Document Management System)')
    print('='*60)
    print(f'  Directory: {BASE_DIR}')
    print(f'  Data库: {DB_PATH}')
    print(f'  默认账户:')
    print(f'    Admin: admin / admin123')
    print(f'='*60)
    print(f'  访问Address: http://localhost:5000')
    print('='*60)

    init_db()
    # （File），For30
    def _delayed_refresh():
        import time
        time.sleep(30)  #
        removed = _do_refresh_files()
        if removed > 0:
            print(f'[启动] File清单Refresh完毕，清理 {removed}  残留Record')
        else:
            print('[启动] File清单Check完毕，无残留Record')
    threading.Thread(target=_delayed_refresh, daemon=True).start()
    print('[启动] File清单CheckWillAt后台执行（30秒后）')

    # threaded=True Can

    # =====================================================================
    # ：3UpdateFull-text index
    # =====================================================================
    from apscheduler.schedulers.background import BackgroundScheduler
    from apscheduler.triggers.cron import CronTrigger

    def _weekly_index_update():
        """每周日3点自动重建Full-text index"""
        scheduler_cfg = get_scheduler_config()
        if not scheduler_cfg.get('weekly_index_enabled'):
            print("[SCHEDULER] 每周Index任务Already跳过（开关未Enable）")
            return
        import subprocess, sys, os
        p = os.path.join(os.path.dirname(__file__), "system", "rebuild_index.py")
        try:
            subprocess.run([sys.executable, p, "--auto"], capture_output=True, timeout=7200)
            print("[SCHEDULER] 每周IndexUpdate完成")
        except Exception as e:
            print(f"[SCHEDULER] 每周IndexUpdateFailed: {e}")

    def _weekly_preview_cleanup():
        """每周日凌晨1点清理Preview Cache中超过7天PDF"""
        scheduler_cfg = get_scheduler_config()
        if not scheduler_cfg.get('weekly_cleanup_enabled'):
            print("[SCHEDULER] Preview清理任务Already跳过（开关未Enable）")
            return
        import time
        preview_dir = os.path.join(os.path.dirname(__file__), 'temp', 'preview')
        if not os.path.isdir(preview_dir):
            return
        now = time.time()
        max_age = 7 * 86400
        cleaned = 0
        try:
            for f in os.listdir(preview_dir):
                fp = os.path.join(preview_dir, f)
                if os.path.isfile(fp) and (now - os.path.getmtime(fp)) > max_age:
                    os.remove(fp)
                    cleaned += 1
            print(f"[SCHEDULER] Preview Cache清理完成，Delete {cleaned} 个过期File")
        except Exception as e:
            print(f"[SCHEDULER] Preview Cache清理Failed: {e}")

    try:
        scheduler = BackgroundScheduler()
        scheduler.add_job(
            _weekly_index_update,
            CronTrigger(day_of_week='sun', hour=3, minute=0),
            id='weekly_index_update',
            name='每周日3点IndexUpdate',
            replace_existing=True
        )
        scheduler.add_job(
            _weekly_preview_cleanup,
            CronTrigger(day_of_week='sun', hour=1, minute=0),
            id='weekly_preview_cleanup',
            name='每周日凌晨1点Preview清理',
            replace_existing=True
        )
        scheduler.start()
        print("[SCHEDULER] 定时任务Already启动（IndexUpdate: 每周日03:00, Preview清理: 每周日01:00）")
    except Exception as e:
        print(f"[SCHEDULER] 定时任务启动Failed: {e}")

    app.run(host='0.0.0.0', port=5000, debug=False, threaded=True)




