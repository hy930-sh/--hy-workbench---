#!/usr/bin/env python3
"""
HY工作台 - 后端服务
基于 Python 内置 http.server，无需额外依赖（仅需 openpyxl / pdfplumber / pypdf）

提供能力：
  1. 静态页面服务（工作台 HTML）
  2. POST /api/convert-excel  上传已有 Excel → 自动列映射 → 生成模板 zip
  3. POST /api/extract-pdf    上传 PDF + 页码范围 → 提取文本
  4. POST /api/generate       提交单词 JSON 数据 → 生成模板 zip

启动：python app.py  （默认端口 8080）
"""

import os
import sys
import io
import json
import re
import zipfile
import tempfile
import shutil
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import parse_qs, unquote

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.join(os.path.dirname(SCRIPT_DIR), ".trae", "skills", "word-excel-organizer")
TEMPLATE_PATH = os.path.join(SKILL_DIR, "导入课程词汇模板.xlsx")

# 复用 generate_word_excel 的核心函数
sys.path.insert(0, SKILL_DIR)
from generate_word_excel import build, format_phonetic, sanitize_xlsx  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

import pdfplumber  # noqa: E402

# OCR 相关（用于扫描件 PDF 和图片识别）
try:
    import pytesseract
    from PIL import Image
    HAS_OCR = True
except ImportError:
    HAS_OCR = False

# 音标相关（CMU 词典 → IPA）
try:
    import pronouncing
    HAS_PHONETIC = True
except ImportError:
    HAS_PHONETIC = False

# CMU (ARPAbet) → IPA 映射表
_CMU_TO_IPA = {
    'AA': 'ɑ', 'AE': 'æ', 'AH': 'ʌ', 'AO': 'ɔ', 'AW': 'aʊ', 'AY': 'aɪ',
    'EH': 'ɛ', 'ER': 'ɝ', 'EY': 'eɪ', 'IH': 'ɪ', 'IY': 'i', 'OW': 'oʊ',
    'OY': 'ɔɪ', 'UH': 'ʊ', 'UW': 'u',
    'B': 'b', 'CH': 'tʃ', 'D': 'd', 'DH': 'ð', 'F': 'f', 'G': 'ɡ',
    'HH': 'h', 'JH': 'dʒ', 'K': 'k', 'L': 'l', 'M': 'm', 'N': 'n',
    'NG': 'ŋ', 'P': 'p', 'R': 'r', 'S': 's', 'SH': 'ʃ', 'T': 't',
    'TH': 'θ', 'V': 'v', 'W': 'w', 'Y': 'j', 'Z': 'z', 'ZH': 'ʒ',
}

def cmu_to_ipa(cmu):
    """将 CMU 音标（如 AE1 P AH0 L）转换为 IPA（如 ˈæp.əl）"""
    if not cmu:
        return ''
    tokens = cmu.strip().split()
    ipa_tokens = []
    for tok in tokens:
        # 分离音素和重音数字
        stress_num = None
        phone = tok
        if tok[-1].isdigit():
            stress_num = int(tok[-1])
            phone = tok[:-1]

        # AH 在非重读音节(0)中是 ə，重读音节中是 ʌ
        if phone == 'AH':
            ipa = 'ə' if stress_num == 0 else 'ʌ'
        # ER 在非重读音节中是 ɚ
        elif phone == 'ER':
            ipa = 'ɚ' if stress_num == 0 else 'ɝ'
        else:
            ipa = _CMU_TO_IPA.get(phone, phone.lower())

        # 重音符号放在音节开头
        stress = ''
        if stress_num == 1:
            stress = 'ˈ'
        elif stress_num == 2:
            stress = 'ˌ'
        ipa_tokens.append(stress + ipa)

    return ''.join(ipa_tokens)


# ============ 列名识别 ============
WORD_KEYS = ["单词", "word", "词汇", "英文", "english", "词语"]
MEANING_KEYS = ["意思", "释义", "中文", "译文", "翻译", "translation", "meaning", "词性及意思", "词义"]
POS_KEYS = ["词性", "pos", "part of speech"]
PHONETIC_KEYS = ["音标", "phonetic", "pronunciation", "发音"]

def detect_columns(headers):
    """根据表头自动识别列映射，返回 {word, meaning, pos, phonetic} 的列索引（0-based），找不到为 None"""
    result = {"word": None, "meaning": None, "pos": None, "phonetic": None}
    for i, h in enumerate(headers):
        h = str(h).strip().lower()
        if result["word"] is None and any(k in h for k in WORD_KEYS):
            result["word"] = i
        if result["phonetic"] is None and any(k in h for k in PHONETIC_KEYS):
            result["phonetic"] = i
        if result["pos"] is None and any(k in h for k in POS_KEYS):
            result["pos"] = i
        if result["meaning"] is None and any(k in h for k in MEANING_KEYS):
            result["meaning"] = i
    return result


def parse_page_range(page_str, total_pages):
    """解析页码范围，如 '1-5,8,10-12'，返回页码列表（1-based）"""
    if not page_str or not page_str.strip():
        return list(range(1, total_pages + 1))
    pages = set()
    for part in page_str.split(","):
        part = part.strip()
        if "-" in part:
            try:
                a, b = part.split("-", 1)
                a, b = int(a), int(b)
                pages.update(range(max(1, a), min(total_pages, b) + 1))
            except ValueError:
                continue
        else:
            try:
                p = int(part)
                if 1 <= p <= total_pages:
                    pages.add(p)
            except ValueError:
                continue
    return sorted(pages)


# ============ HTTP 请求处理器 ============
class Handler(BaseHTTPRequestHandler):

    def log_message(self, format, *args):
        print(f"[{self.log_date_time_string()}] {args[0]}")

    # ---------- 静态文件 ----------
    def do_GET(self):
        path = self.path.split("?")[0]
        if path == "/" or path == "/index.html":
            self._serve_file(os.path.join(SCRIPT_DIR, "index.html"), "text/html; charset=utf-8")
        elif path == "/manifest.json":
            self._serve_file(os.path.join(SCRIPT_DIR, "manifest.json"), "application/json; charset=utf-8")
        elif path.startswith("/footprint/"):
            # 足迹照片访问
            filename = os.path.basename(path)
            footprint_dir = os.path.join(SCRIPT_DIR, "footprint")
            filepath = os.path.join(footprint_dir, filename)
            mime = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
                    ".gif": "image/gif", ".webp": "image/webp"}
            ext = os.path.splitext(filename)[1].lower()
            self._serve_file(filepath, mime.get(ext, "application/octet-stream"))
        elif path.lower().endswith((".pdf", ".png", ".jpg", ".jpeg", ".gif", ".webp", ".xlsx", ".xls", ".csv")):
            # 支持静态文件访问（测试/预览用）
            filename = os.path.basename(path)
            filepath = os.path.join(SCRIPT_DIR, filename)
            if os.path.exists(filepath):
                mime = {
                    ".pdf": "application/pdf",
                    ".png": "image/png",
                    ".jpg": "image/jpeg",
                    ".jpeg": "image/jpeg",
                    ".gif": "image/gif",
                    ".webp": "image/webp",
                    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ".xls": "application/vnd.ms-excel",
                    ".csv": "text/csv",
                }
                ext = os.path.splitext(filename)[1].lower()
                self._serve_file(filepath, mime.get(ext, "application/octet-stream"))
            else:
                self.send_error(404)
        else:
            self.send_error(404)

    def _serve_file(self, filepath, content_type):
        if not os.path.exists(filepath):
            self.send_error(404)
            return
        with open(filepath, "rb") as f:
            data = f.read()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    # ---------- API ----------
    def do_POST(self):
        path = self.path.split("?")[0]
        if path == "/api/convert-excel":
            self._handle_convert_excel()
        elif path == "/api/extract-pdf":
            self._handle_extract_pdf()
        elif path == "/api/generate":
            self._handle_generate()
        elif path == "/api/suggest-name":
            self._handle_suggest_name()
        elif path == "/api/ocr-image":
            self._handle_ocr_image()
        elif path == "/api/phonetic":
            self._handle_phonetic()
        elif path == "/api/upload-footprint":
            self._handle_upload_footprint()
        else:
            self.send_error(404)

    def _read_multipart(self):
        """解析 multipart/form-data，返回 (fields_dict, files_dict)
        files_dict: {field_name: (filename, content_bytes)}"""
        content_type = self.headers.get("Content-Type", "")
        if "multipart/form-data" not in content_type:
            return {}, {}
        boundary = content_type.split("boundary=")[-1].strip().encode()
        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length)

        fields = {}
        files = {}
        parts = body.split(b"--" + boundary)
        for part in parts:
            if not part or part in (b"--", b"--\r\n", b"\r\n"):
                continue
            part = part.strip(b"\r\n")
            if b"\r\n\r\n" not in part:
                continue
            header, content = part.split(b"\r\n\r\n", 1)
            content = content.rstrip(b"\r\n")
            header_str = header.decode("utf-8", errors="replace")
            name_match = re.search(r'name="([^"]*)"', header_str)
            if not name_match:
                continue
            name = name_match.group(1)
            filename_match = re.search(r'filename="([^"]*)"', header_str)
            if filename_match:
                files[name] = (filename_match.group(1), content)
            else:
                fields[name] = content.decode("utf-8", errors="replace")
        return fields, files

    def _json_response(self, data, code=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _zip_response(self, zip_bytes, filename):
        # HTTP header 只支持 latin-1，中文文件名需 percent-encode（RFC 5987）
        from urllib.parse import quote
        encoded = quote(filename)
        self.send_response(200)
        self.send_header("Content-Type", "application/zip")
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{encoded}")
        self.send_header("Content-Length", str(len(zip_bytes)))
        self.end_headers()
        self.wfile.write(zip_bytes)

    # ---------- 足迹照片上传 ----------
    def _handle_upload_footprint(self):
        fields, files = self._read_multipart()
        if "file" not in files:
            self._json_response({"error": "请上传照片"}, 400)
            return
        filename, content = files["file"]
        ext = os.path.splitext(filename)[1].lower()
        if ext not in (".png", ".jpg", ".jpeg", ".gif", ".webp"):
            self._json_response({"error": "仅支持图片格式"}, 400)
            return
        import time, uuid
        footprint_dir = os.path.join(SCRIPT_DIR, "footprint")
        os.makedirs(footprint_dir, exist_ok=True)
        save_name = f"{int(time.time())}_{uuid.uuid4().hex[:8]}{ext}"
        save_path = os.path.join(footprint_dir, save_name)
        with open(save_path, "wb") as f:
            f.write(content)
        self._json_response({"url": f"/footprint/{save_name}"})

    # ---------- Excel 转模板 ----------
    def _handle_convert_excel(self):
        fields, files = self._read_multipart()
        if "file" not in files:
            self._json_response({"error": "请上传 Excel 文件"}, 400)
            return
        filename, content = files["file"]
        output_name = fields.get("output", os.path.splitext(filename)[0])
        is_xls = filename.lower().endswith(".xls")

        try:
            rows = []
            if is_xls:
                # 旧版 .xls 格式用 xlrd 读取
                import xlrd
                wb = xlrd.open_workbook(file_contents=content)
                ws = wb.sheet_by_index(0)
                for r in range(ws.nrows):
                    rows.append([ws.cell_value(r, c) for c in range(ws.ncols)])
            else:
                wb = load_workbook(io.BytesIO(content), data_only=True)
                ws = wb.active
                rows = list(ws.iter_rows(values_only=True))
            if not rows:
                self._json_response({"error": "Excel 为空"}, 400)
                return

            headers = [str(c).strip() if c is not None else "" for c in rows[0]]
            mapping = detect_columns(headers)

            if mapping["word"] is None:
                # 没识别到表头，尝试按列顺序推断：第1列单词，第2列意思，第3列音标
                mapping = {"word": 0, "meaning": 1 if len(headers) > 1 else None,
                           "pos": None, "phonetic": 2 if len(headers) > 2 else None}

            word_data = []
            for row in rows[1:]:
                if not row or all(c is None for c in row):
                    continue
                word = str(row[mapping["word"]]).strip() if mapping["word"] is not None and len(row) > mapping["word"] else ""
                if not word:
                    continue
                # 合并词性和意思
                pos = ""
                meaning = ""
                if mapping["pos"] is not None and len(row) > mapping["pos"] and row[mapping["pos"]]:
                    pos = str(row[mapping["pos"]]).strip()
                if mapping["meaning"] is not None and len(row) > mapping["meaning"] and row[mapping["meaning"]]:
                    meaning = str(row[mapping["meaning"]]).strip()
                if pos and meaning:
                    if pos.endswith('.'):
                        pos_meaning = f"{pos} {meaning}"
                    else:
                        pos_meaning = f"{pos}. {meaning}"
                elif pos:
                    pos_meaning = pos
                else:
                    pos_meaning = meaning
                # 音标
                phonetic = ""
                if mapping["phonetic"] is not None and len(row) > mapping["phonetic"] and row[mapping["phonetic"]]:
                    phonetic = str(row[mapping["phonetic"]]).strip()

                word_data.append({
                    "word": word,
                    "pos_meaning": pos_meaning,
                    "phonetic": phonetic
                })

            # 生成 zip
            tmp_dir = tempfile.mkdtemp()
            try:
                base = os.path.join(tmp_dir, output_name)
                build(TEMPLATE_PATH, word_data, base)
                zip_path = base + ".zip"
                with open(zip_path, "rb") as f:
                    zip_bytes = f.read()
            finally:
                shutil.rmtree(tmp_dir, ignore_errors=True)

            self._zip_response(zip_bytes, f"{output_name}.zip")
        except Exception as e:
            self._json_response({"error": f"处理失败：{str(e)}"}, 500)

    # ---------- PDF 文本提取 ----------
    def _handle_extract_pdf(self):
        fields, files = self._read_multipart()
        if "file" not in files:
            self._json_response({"error": "请上传 PDF 文件"}, 400)
            return
        filename, content = files["file"]
        page_range = fields.get("pages", "")

        try:
            with pdfplumber.open(io.BytesIO(content)) as pdf:
                total = len(pdf.pages)
                pages = parse_page_range(page_range, total)
                text_parts = []
                for p in pages:
                    page = pdf.pages[p - 1]
                    t = page.extract_text() or ""
                    text_parts.append(f"--- 第 {p} 页 ---\n{t}")
                full_text = "\n".join(text_parts)

            # 如果提取文本过少，尝试 OCR（扫描件 PDF）
            if len(full_text.strip()) < 20 and HAS_OCR:
                try:
                    from pdf2image import convert_from_bytes
                    ocr_parts = []
                    for p in pages:
                        imgs = convert_from_bytes(content, first_page=p, last_page=p, dpi=200)
                        if imgs:
                            ocr_text = pytesseract.image_to_string(imgs[0], lang='eng+chi_sim')
                            ocr_parts.append(f"--- 第 {p} 页（OCR）---\n{ocr_text}")
                    if ocr_parts:
                        full_text = "\n".join(ocr_parts)
                except Exception as oe:
                    pass  # OCR 失败时保持原文本

            self._json_response({
                "total_pages": total,
                "extracted_pages": pages,
                "text": full_text
            })
        except Exception as e:
            self._json_response({"error": f"PDF 提取失败：{str(e)}"}, 500)

    # ---------- 单词数据生成 ----------
    def _handle_generate(self):
        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length).decode("utf-8")
        try:
            data = json.loads(body)
            rows = data.get("words", [])
            output_name = data.get("output", "单词表")
            if not rows:
                self._json_response({"error": "单词数据为空"}, 400)
                return
            tmp_dir = tempfile.mkdtemp()
            try:
                base = os.path.join(tmp_dir, output_name)
                build(TEMPLATE_PATH, rows, base)
                zip_path = base + ".zip"
                with open(zip_path, "rb") as f:
                    zip_bytes = f.read()
            finally:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            self._zip_response(zip_bytes, f"{output_name}.zip")
        except json.JSONDecodeError:
            self._json_response({"error": "JSON 格式错误"}, 400)
        except Exception as e:
            self._json_response({"error": f"生成失败：{str(e)}"}, 500)

    # ---------- 智能命名（从 PDF 封面提取出版社+年级册/必修选修） ----------
    def _handle_suggest_name(self):
        fields, files = self._read_multipart()
        if "file" not in files:
            self._json_response({"error": "请上传 PDF 文件"}, 400)
            return
        filename, content = files["file"]
        try:
            # 提取封面（第1页）、版权页（第2页）、封底/后记（最后2页）文本
            with pdfplumber.open(io.BytesIO(content)) as pdf:
                pages = [pdf.pages[0]] if len(pdf.pages) > 0 else []
                if len(pdf.pages) > 1:
                    pages.append(pdf.pages[1])
                # 最后两页（出版社信息常在此）
                if len(pdf.pages) > 2:
                    pages.append(pdf.pages[-1])
                if len(pdf.pages) > 3:
                    pages.append(pdf.pages[-2])
                text = "\n".join((p.extract_text() or "") for p in pages)
                cover_text = pages[0].extract_text() or "" if pages else ""
            suggested = suggest_filename(text, filename)
            self._json_response({"suggested_name": suggested, "cover_text": cover_text[:500]})
        except Exception as e:
            # 提取失败时用原文件名兜底
            base = os.path.splitext(filename)[0]
            self._json_response({"suggested_name": base, "error": str(e)})

    # ---------- 图片 OCR 识别（看图识词用） ----------
    def _handle_ocr_image(self):
        if not HAS_OCR:
            self._json_response({"error": "服务器未安装 OCR 组件（pytesseract/Pillow）"}, 500)
            return
        fields, files = self._read_multipart()
        if "file" not in files:
            self._json_response({"error": "请上传图片文件"}, 400)
            return
        filename, content = files["file"]
        try:
            img = Image.open(io.BytesIO(content))
            # 转为 RGB（兼容 RGBA/灰度）
            if img.mode != "RGB":
                img = img.convert("RGB")
            # OCR 识别（中英文）
            text = pytesseract.image_to_string(img, lang="eng+chi_sim", config="--psm 3")
            # 提取英文单词（去重、过滤、排序）
            raw_words = re.findall(r"[a-zA-Z][a-zA-Z'-]*", text)
            words = sorted(set(w.lower() for w in raw_words if len(w) > 1))
            self._json_response({
                "raw_text": text,
                "words": words,
                "count": len(words)
            })
        except Exception as e:
            self._json_response({"error": f"OCR 识别失败：{str(e)}"}, 500)

    # ---------- 批量获取美式音标（CMU → IPA） ----------
    def _handle_phonetic(self):
        if not HAS_PHONETIC:
            self._json_response({"error": "服务器未安装音标组件（pronouncing）"}, 500)
            return
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(length) if length > 0 else b"{}"
            data = json.loads(body.decode("utf-8"))
            words = data.get("words", [])
            if not isinstance(words, list):
                self._json_response({"error": "words 必须是数组"}, 400)
                return

            result = {}
            for w in words:
                w = str(w).strip().lower()
                if not w:
                    continue
                try:
                    phones = pronouncing.phones_for_word(w)
                    if phones:
                        result[w] = cmu_to_ipa(phones[0])
                    else:
                        result[w] = ""
                except Exception:
                    result[w] = ""

            self._json_response({"phonetics": result})
        except Exception as e:
            self._json_response({"error": f"音标获取失败：{str(e)}"}, 500)


# ============ 智能命名辅助函数 ============
PUBLISHER_PATTERNS = [
    (r"人民教育出版社", "人教"),
    (r"北京师范大学出版社", "北师大"),
    (r"外语教学与研究出版社", "外研"),
    (r"上海教育出版社", "沪教"),
    (r"译林出版社", "译林"),
    (r"福建教育出版社|闽教出版|闽教", "福建教育"),
    (r"河北教育出版社", "冀教"),
    (r"湖南教育出版社", "湘教"),
    (r"山东教育出版社", "鲁教"),
    (r"江苏教育出版社", "苏教"),
    (r"广东教育出版社", "粤教"),
    (r"浙江教育出版社", "浙教"),
    (r"四川教育出版社", "川教"),
    (r"教育科学出版社", "教科"),
    (r"清华大学出版社", "清华"),
    (r"北京大学出版社", "北大"),
]

GRADE_PATTERNS = [
    # 小学/初中
    (r"三[年级]*\s*上(?:册)?", "三上"),
    (r"三[年级]*\s*下(?:册)?", "三下"),
    (r"四[年级]*\s*上(?:册)?", "四上"),
    (r"四[年级]*\s*下(?:册)?", "四下"),
    (r"五[年级]*\s*上(?:册)?", "五上"),
    (r"五[年级]*\s*下(?:册)?", "五下"),
    (r"六[年级]*\s*上(?:册)?", "六上"),
    (r"六[年级]*\s*下(?:册)?", "六下"),
    (r"七[年级]*\s*上(?:册)?", "七上"),
    (r"七[年级]*\s*下(?:册)?", "七下"),
    (r"八[年级]*\s*上(?:册)?", "八上"),
    (r"八[年级]*\s*下(?:册)?", "八下"),
    (r"九[年级]*\s*上(?:册)?", "九上"),
    (r"九[年级]*\s*下(?:册)?", "九下"),
    (r"九[年级]*\s*全(?:册)?", "九全"),
    # 高中必修
    (r"必修\s*第?\s*[一二三四五12345]\s*册?", "必修一"),  # 后续替换具体数字
    (r"必修\s*[一二三四五12345]", "必修一"),
    # 高中选修
    (r"选修\s*第?\s*[一二三四五六123456]\s*册?", "选修一"),
    (r"选修\s*[一二三四五六123456]", "选修一"),
]

CN_NUM = {"一": "一", "二": "二", "三": "三", "四": "四", "五": "五", "六": "六",
          "1": "一", "2": "二", "3": "三", "4": "四", "5": "五", "6": "六"}


def suggest_filename(text, original_filename):
    """从 PDF 文本中提取出版社和年级册/必修选修，生成建议文件名"""
    text = text or ""

    # 1. 识别出版社/教材版本
    publisher = None
    for pattern, short in PUBLISHER_PATTERNS:
        if re.search(pattern, text):
            publisher = short
            break

    # 2. 识别年级册 / 必修 / 选修
    grade = None
    # 先匹配必修/选修（更具体）
    bx_match = re.search(r"必修\s*第?\s*([一二三四五12345])", text)
    if bx_match:
        num = CN_NUM.get(bx_match.group(1), bx_match.group(1))
        grade = f"必修{num}"
    else:
        xx_match = re.search(r"选修\s*第?\s*([一二三四五六123456])", text)
        if xx_match:
            num = CN_NUM.get(xx_match.group(1), xx_match.group(1))
            grade = f"选修{num}"

    # 再匹配年级册
    if not grade:
        for pattern, g in GRADE_PATTERNS[:18]:  # 只匹配年级册的
            m = re.search(pattern, text)
            if m:
                grade = g
                break

    # 3. 组合命名
    if publisher and grade:
        return f"{publisher}-{grade}"
    elif grade:
        return grade
    else:
        # 兜底：用原文件名
        return os.path.splitext(original_filename)[0]


def main():
    import socket
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8080
    # 获取本机局域网 IP
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        local_ip = s.getsockname()[0]
        s.close()
    except Exception:
        local_ip = "127.0.0.1"

    server = HTTPServer(("0.0.0.0", port), Handler)
    print(f"============================================")
    print(f"  HY的工作台已启动")
    print(f"============================================")
    print(f"  本机访问:   http://localhost:{port}")
    print(f"  手机访问:   http://{local_ip}:{port}  (同一WiFi下)")
    print(f"  模板路径:   {TEMPLATE_PATH}")
    print(f"  按 Ctrl+C 停止")
    print(f"============================================")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止")
        server.server_close()


if __name__ == "__main__":
    main()
