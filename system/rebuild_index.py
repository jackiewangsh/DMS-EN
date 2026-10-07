#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
重建 file_index 和 keyword_index
只索引发布文件库（uploads/published/）中的文件
同步策略：以发布文件库为准，不存在的文件删除索引
"""
import sys, os, re, uuid, shutil, glob, time
sys.path.insert(0, os.path.dirname(__file__))
os.chdir(os.path.dirname(__file__))

import pymysql
import pdfplumber
import docx
import jieba

jieba.setLogLevel(jieba.logging.INFO)

STOP = {'the','a','an','and','or','but','in','on','at','to','for','of','with','by',
        'is','are','was','were','be','been','being','have','has','had','do','does',
        'did','will','would','could','should','may','might','must','shall','can',
        'this','that','these','those','i','you','he','she','it','we','they','what',
        'which','who','whom','if','then','else','when','where','why','how','all',
        'each','every','both','few','more','most','other','some','such','no','nor',
        'not','only','own','same','so','than','too','very','just','also','now'}

# 发布文件库根目录（从 system/ 目录向上一级）
PUBLISHED_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'uploads', 'published')
PROGRESS_FILE = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data', 'rebuild_progress.json')

def write_progress(indexed, total, status, message):
    """写入进度到 JSON 文件"""
    import json
    data = {
        'indexed': indexed,
        'total': total,
        'status': status,  # running, done, error
        'message': message,
        'updated_at': time.strftime('%Y-%m-%d %H:%M:%S')
    }
    try:
        with open(PROGRESS_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f'[WARN] 写入进度失败: {e}')

def get_content(path, ext):
    ext = ext.lower()
    text = ''
    try:
        if ext == '.pdf':
            with pdfplumber.open(path) as pdf:
                for page in pdf.pages:
                    t = page.extract_text() or ''
                    text += t + '\n'
                    # 提取表格文字
                    tables = page.extract_tables()
                    for table in tables:
                        for row in table:
                            for cell in row:
                                if cell:
                                    text += str(cell) + ' '
                        text += '\n'
        elif ext == '.docx':
            doc = docx.Document(path)
            for para in doc.paragraphs:
                text += para.text + '\n'
            for tbl in doc.tables:
                for row in tbl.rows:
                    for cell in row.cells:
                        if cell.text.strip():
                            text += cell.text.strip() + '\n'
        elif ext == '.doc':
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
            except Exception as e:
                print(f'  [ERROR] doc extract: {e}')
        elif ext == '.xlsx':
            import openpyxl
            wb = openpyxl.load_workbook(path, data_only=True)
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                for row in ws.iter_rows(values_only=True):
                    for cell in row:
                        if cell:
                            text += str(cell) + ' '
                    text += '\n'
        elif ext == '.xls':
            import xlrd
            wb = xlrd.open_workbook(path)
            for sheet in wb.sheets():
                for row_idx in range(sheet.nrows):
                    for col_idx in range(sheet.ncols):
                        val = sheet.cell_value(row_idx, col_idx)
                        if val:
                            text += str(val) + ' '
                    text += '\n'
        elif ext == '.pptx':
            from pptx import Presentation
            prs = Presentation(path)
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, 'text'):
                        text += shape.text + '\n'
        elif ext == '.txt':
            with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                text = f.read()
    except Exception as e:
        print(f'  [ERROR] extract {path}: {e}')
    return text

def tokenize(content):
    try:
        # 去除中文字符之间的空格（如 "王 建国" → "王建国"）
        for _ in range(3):
            content = re.sub(r'([\u4e00-\u9fa5])\s+([\u4e00-\u9fa5])', r'\1\2', content)
        words = jieba.cut_for_search(content)
        result = []
        for w in words:
            w = w.strip().lower()
            if len(w) < 2 or w in STOP:
                continue
            if re.match(r'^[\u4e00-\u9fff]$', w):
                continue
            result.append(w)
        return result
    except Exception:
        return []

def scan_published_files():
    """扫描发布文件库，返回 {filename: full_path}"""
    published_files = {}
    if not os.path.exists(PUBLISHED_DIR):
        print(f'[WARN] 发布文件库不存在: {PUBLISHED_DIR}')
        return published_files
    
    for root, dirs, files in os.walk(PUBLISHED_DIR):
        for f in files:
            if f.startswith('.'):
                continue
            full_path = os.path.join(root, f)
            # 使用文件名作为 key（与数据库 filename 字段一致）
            published_files[f] = full_path
    
    return published_files

def rebuild_all():
    # 写入初始进度
    write_progress(0, 0, 'running', '开始扫描文件...')
    
    conn = pymysql.connect(host='localhost', user='dms_user', password='DmsUser123!', database='dms', charset='utf8mb4')
    cur = conn.cursor()
    
    print('[INFO] 扫描发布文件库...')
    published_files = scan_published_files()
    print(f'[INFO] 发布文件库共有 {len(published_files)} 个文件')
    
    # 第一步：删除不在发布文件库中的索引
    print('[INFO] 检查并删除无效索引...')
    
    # 获取所有索引中的 file_id
    cur.execute('SELECT file_id FROM file_index')
    indexed_ids = set(row[0] for row in cur.fetchall())
    
    if indexed_ids:
        # 找出需要删除的索引（索引中存在但发布文件库中没有的文件）
        to_delete = indexed_ids - set(published_files.keys())
        
        if to_delete:
            print(f'[INFO] 删除 {len(to_delete)} 个无效索引...')
            for fid in to_delete:
                cur.execute('DELETE FROM file_index WHERE file_id=%s', (fid,))
                cur.execute('DELETE FROM keyword_index WHERE file_id=%s', (fid,))
            conn.commit()
            print(f'[INFO] 已删除 {len(to_delete)} 个无效索引')
        else:
            print('[INFO] 没有需要删除的索引')
    
    # 第二步：从数据库查询发布文件库中实际存在的文件记录
    print('[INFO] 查询数据库中的文件记录...')
    
    # 构建发布文件库的文件名列表（用于 SQL 查询）
    if published_files:
        file_list = list(published_files.keys())
        # 分批查询，每批 100 个
        placeholders = ','.join(['%s'] * len(file_list))
        cur.execute(f"""
            SELECT id, filename, original_name, file_path, file_size, status
            FROM files
            WHERE filename IN ({placeholders})
            AND status != '作废'
        """, file_list)
    else:
        cur.execute("""
            SELECT id, filename, original_name, file_path, file_size, status
            FROM files
            WHERE 1=0
        """)
    
    rows = cur.fetchall()
    print(f'[INFO] 数据库中找到 {len(rows)} 个待索引文件')
    
    indexed = 0
    errors = 0
    skipped = 0
    
    for row in rows:
        files_id, filename, original_name, file_path, file_size, status = row
        
        # 检查文件是否在发布文件库中
        if filename not in published_files:
            print(f'  [SKIP] 文件不在发布文件库: {filename}')
            skipped += 1
            continue
        
        # 使用发布文件库中的实际路径
        actual_path = published_files[filename]
        
        if not os.path.exists(actual_path):
            print(f'  [SKIP] 文件实际不存在: {actual_path}')
            errors += 1
            continue
        
        ext = os.path.splitext(original_name)[1] if original_name else os.path.splitext(filename)[1]
        content = get_content(actual_path, ext)
        if not content:
            print(f'  [WARN] 无法提取内容: {filename}')
        
        # 先删除旧索引（如果存在）
        cur.execute('DELETE FROM file_index WHERE file_id=%s', (filename,))
        cur.execute('DELETE FROM keyword_index WHERE file_id=%s', (filename,))
        
        # 插入新索引
        content_snippet = content[:1000000]
        
        try:
            cur.execute("""
                INSERT INTO file_index (file_id, file_path, original_name, content, file_ext, file_size, indexed_at)
                VALUES (%s,%s,%s,%s,%s,%s,NOW())
            """, (filename, actual_path, original_name, content_snippet, ext, file_size or 0))
            
            # 关键词
            tokens = tokenize(content)
            tf = {}
            for t in tokens:
                tf[t] = tf.get(t, 0) + 1
            for term, freq in tf.items():
                cur.execute("""
                    INSERT INTO keyword_index (term, file_id, frequency, indexed_at)
                    VALUES (%s,%s,%s,NOW())
                """, (term, filename, freq))
            
            conn.commit()
            indexed += 1
            if indexed % 20 == 0:
                print(f'  已索引 {indexed}/{len(rows)}...')
                write_progress(indexed, len(rows), 'running', f'已索引 {indexed}/{len(rows)}...')
        except Exception as e:
            conn.rollback()
            print(f'  [ERROR] DB: {e}')
            errors += 1
    
    # 第三步：显示最终统计
    cur.execute('SELECT COUNT(*) FROM file_index')
    total_indexed = cur.fetchone()[0]
    
    print(f'\n[DONE] 索引重建完成:')
    print(f'  - 成功索引: {indexed}')
    print(f'  - 跳过: {skipped}')
    print(f'  - 失败: {errors}')
    print(f'  - 总计索引: {total_indexed}')
    
    # 写入最终进度
    write_progress(indexed, indexed, 'done', f'索引重建完成: 成功{indexed}, 跳过{skipped}, 失败{errors}, 总计{total_indexed}')
    
    conn.close()

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('--auto', action='store_true', help='API调用模式')
    args = parser.parse_args()

    print('=' * 60)
    print('  DMS 全文索引重建（以发布文件库为准）')
    print('=' * 60)
    print(f'发布文件库: {PUBLISHED_DIR}')
    print('=' * 60)
    rebuild_all()