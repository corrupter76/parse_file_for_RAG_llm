from flask import Flask, request, jsonify, Response
from werkzeug.utils import secure_filename
import os
os.environ["PROTOCOL_BUFFERS_PYTHON_IMPLEMENTATION"] = "python"

import pandas as pd
import shutil
import logging
import numpy as np
import requests
import base64
import json
import io
import mimetypes
import ftfy
import time
import traceback
from pathlib import Path
import subprocess
import re
import html
from datetime import datetime
import os
# Set environment variable to avoid OpenMP conflicts
os.environ['KMP_DUPLICATE_LIB_OK'] = 'TRUE'
import cv2
import easyocr
import pdf2image
from collections import defaultdict

import logging
import os
import os
import psutil
import logging
from pdfplumber import open as pdfplumber_open
from pdf2image import convert_from_path
from PIL import Image
import pytesseract

log_file_path = os.path.abspath('file_parser2.log')  # 获取绝对路径
print(f"日志文件路径：{log_file_path}")

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(log_file_path)
    ]
)

logger = logging.getLogger(__name__)

# 初始化 mimetypes
mimetypes.init()

# Flask 应用
app = Flask(__name__)

# 设置最大内容长度为500MB
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024

# 上传文件存储目录
UPLOAD_FOLDER = Path("uploads")
UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
app.config["UPLOAD_FOLDER"] = str(UPLOAD_FOLDER)

# 允许的文件扩展名
ALLOWED_EXTENSIONS = {
    "pdf", "json", "csv", "docx", "xlsx", "md", "pptx", "html", "txt", "doc", "ppt",
    "png", "jpg", "jpeg", "xls", "gif", "bmp", "tiff"
}

def convert_to_pdf(file_path, original_filename=None):
    """
    将各种文件格式直接转换为PDF，特别处理PPT文件的命名问题
    
    Args:
        file_path (str): 文件系统安全的文件路径
        original_filename (str, optional): 原始文件名，如果提供则用来命名PDF文件
        
    Returns:
        dict: 包含成功状态和PDF路径的结果
    """
    try:
        # 可能的LibreOffice路径
        libreoffice_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            r"D:\LibreOffice\program\soffice.exe",
            "/usr/bin/soffice",
            "/usr/bin/libreoffice",
            "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        ]
        
        soffice_path = None
        for path in libreoffice_paths:
            if os.path.exists(path):
                soffice_path = path
                break
        
        if not soffice_path:
            logger.warning("未找到 LibreOffice 路径")
            return {"success": False, "error": "未找到 LibreOffice 路径"}
        
        # 提取文件信息
        file_basename = os.path.basename(file_path)
        file_ext = os.path.splitext(file_basename)[1].lower()
        abs_file_path = os.path.abspath(str(file_path))
        abs_output_dir = os.path.dirname(abs_file_path)
        is_ppt = file_ext in ['.ppt', '.pptx']
        
        logger.info(f"文件类型: {file_ext}, 是否是PPT文件: {is_ppt}")
        
        # 计算目标PDF文件名
        if original_filename:
            original_name_without_ext = os.path.splitext(original_filename)[0]
            pdf_filename = f"{original_name_without_ext}.pdf"
            # 生成文件系统安全的路径
            # safe_pdf_filename = secure_filename(pdf_filename)
            # target_pdf_path = os.path.join(abs_output_dir, safe_pdf_filename)

            target_pdf_path = os.path.join(abs_output_dir, pdf_filename)

        else:
            # 如果没有原始文件名，使用上传后的安全文件名
            file_name_without_ext = os.path.splitext(file_basename)[0]
            target_pdf_path = os.path.join(abs_output_dir, f"{file_name_without_ext}.pdf")

        logger.info(f"目标PDF路径: {target_pdf_path}")
        
        # 创建临时目录用于转换，解决PPT文件命名问题
        if is_ppt:
            temp_dir = os.path.join(abs_output_dir, "temp_convert_" + str(int(time.time())))
            os.makedirs(temp_dir, exist_ok=True)
            logger.info(f"为PPT创建临时目录: {temp_dir}")
            output_dir = temp_dir
        else:
            output_dir = abs_output_dir
        
        # 执行LibreOffice转换命令
        logger.info(f"尝试将文件直接转换为PDF: {file_path}")
        
        cmd = [soffice_path, "--headless", "--convert-to", "pdf", abs_file_path, "--outdir", output_dir]
        logger.info(f"执行命令: {' '.join(cmd)}")
        
        process = subprocess.run(
            cmd, 
            capture_output=True, 
            text=True, 
            check=False,
            timeout=120
        )
        
        # 处理LibreOffice输出
        if process.returncode != 0:
            logger.error(f"LibreOffice转换失败，返回码: {process.returncode}")
            logger.error(f"错误信息: {process.stderr}")
        
        # 查找生成的PDF文件
        if is_ppt:
            # 对于PPT，查找临时目录中的所有PDF文件
            pdf_files = [f for f in os.listdir(temp_dir) if f.endswith('.pdf')]
            if pdf_files:
                logger.info(f"在临时目录中找到PDF文件: {pdf_files}")
                temp_pdf_path = os.path.join(temp_dir, pdf_files[0])
                
                # 移动到目标位置并使用原始文件名
                try:
                    shutil.move(temp_pdf_path, target_pdf_path)
                    logger.info(f"将PPT生成的PDF从 {temp_pdf_path} 移动到 {target_pdf_path}")
                    final_pdf_path = target_pdf_path
                except Exception as e:
                    logger.error(f"移动PDF失败: {str(e)}")
                    final_pdf_path = temp_pdf_path
            else:
                logger.error(f"在临时目录中未找到PDF文件: {temp_dir}")
                return {"success": False, "error": "PPT转PDF失败，未找到输出文件"}
                
            # 清理临时目录
            try:
                if os.path.exists(temp_dir) and temp_dir != abs_output_dir:
                    shutil.rmtree(temp_dir)
                    logger.info(f"已清理临时目录: {temp_dir}")
            except Exception as e:
                logger.warning(f"清理临时目录失败: {str(e)}")
        else:
            # 非PPT文件的正常处理流程
            expected_pdf = os.path.splitext(file_basename)[0] + ".pdf"
            expected_pdf_path = os.path.join(abs_output_dir, expected_pdf)
            
            if os.path.exists(expected_pdf_path):
                if expected_pdf_path != target_pdf_path and original_filename:
                    try:
                        shutil.move(expected_pdf_path, target_pdf_path)
                        logger.info(f"已将PDF从 {expected_pdf_path} 重命名为 {target_pdf_path}")
                        final_pdf_path = target_pdf_path
                    except Exception as rename_err:
                        logger.warning(f"重命名PDF文件失败: {str(rename_err)}")
                        final_pdf_path = expected_pdf_path
                else:
                    final_pdf_path = expected_pdf_path
            else:
                # 查找最近创建的PDF文件
                all_files = os.listdir(abs_output_dir)
                recent_pdfs = [f for f in all_files if f.endswith('.pdf') and 
                              os.path.getmtime(os.path.join(abs_output_dir, f)) > 
                              (time.time() - 30)]
                
                if recent_pdfs:
                    found_pdf = os.path.join(abs_output_dir, recent_pdfs[0])
                    logger.info(f"找到可能的PDF文件: {found_pdf}")
                    
                    if original_filename and found_pdf != target_pdf_path:
                        try:
                            shutil.move(found_pdf, target_pdf_path)
                            logger.info(f"已将找到的PDF从 {found_pdf} 重命名为 {target_pdf_path}")
                            final_pdf_path = target_pdf_path
                        except Exception as rename_err:
                            logger.warning(f"重命名找到的PDF文件失败: {str(rename_err)}")
                            final_pdf_path = found_pdf
                    else:
                        final_pdf_path = found_pdf
                else:
                    error_msg = process.stderr if process.stderr else "未知错误"
                    logger.error(f"PDF转换失败，未找到输出文件。错误: {error_msg}")
                    return {"success": False, "error": f"转换失败: {error_msg}"}
        
        # 确认最终PDF是否存在
        if os.path.exists(final_pdf_path):
            logger.info(f"PDF转换成功: {final_pdf_path}")
            return {"success": True, "pdf_path": final_pdf_path}
        else:
            logger.error(f"最终PDF文件不存在: {final_pdf_path}")
            return {"success": False, "error": "转换后的PDF文件不存在"}
    
    except Exception as e:
        logger.error(f"PDF转换失败: {str(e)}")
        logger.error(traceback.format_exc())
        return {"success": False, "error": str(e)}



def save_parsed_data_as_markdown(parsed_data, file_path, original_filename=None, userId=None):
    """
    Save parsed data as Markdown file with specified structure
    
    Args:
        parsed_data (dict): Parsed data to save
        file_path (str): Filesystem-safe file path
        original_filename (str, optional): Original filename for display
        userId (str, optional): User ID for directory structure
        
    Returns:
        dict: Result with Markdown file path and status
    """
    try:
        # Ensure we have the original filename
        if not original_filename and "file_info" in parsed_data and "original_filename" in parsed_data["file_info"]:
            original_filename = parsed_data["file_info"]["original_filename"]
            logger.info(f"Got original filename from parsed data: {original_filename}")
        
        # Generate the markdown filename
        if original_filename:
            original_name_without_ext = os.path.splitext(original_filename)[0]
            logger.info(f"Using original filename for Markdown: {original_name_without_ext}")
            md_file_name = f"{original_name_without_ext}.md"
        else:
            file_name_without_ext = os.path.splitext(os.path.basename(file_path))[0]
            logger.info(f"Using current filename for Markdown: {file_name_without_ext}")
            md_file_name = f"{file_name_without_ext}.md"
            original_name_without_ext = file_name_without_ext
        
        # Create the directory path based on userId, similar to the JSON example
        if userId:
            save_dir = f"upload_files_json_all/{userId}"
            if not os.path.exists(save_dir):
                os.makedirs(save_dir, exist_ok=True)
            md_file_path = os.path.join(save_dir, md_file_name)
        else:
            # Fallback to original behavior if no userId is provided
            md_file_path = os.path.join(os.path.dirname(file_path), md_file_name)
        
        logger.info(f"Saving Markdown to: {md_file_path}")
        
        # Modify file_info to use original filename for display
        file_info = dict(parsed_data.get("file_info", {}))
        file_info["filename"] = original_filename if original_filename else os.path.basename(file_path)
        
        # Convert parsed data to Markdown, passing file_info with original filename
        md_content = convert_to_markdown(parsed_data, file_info)
        
        # Create Markdown content with specified structure
        structured_md = f"""{{
   '名称': '{original_filename if original_filename else os.path.basename(file_path)}',
   'content': '''
{md_content}
'''
}}"""
        
        # Save as Markdown file
        with open(md_file_path, 'w', encoding='utf-8') as f:
            f.write(structured_md)
            
        logger.info(f"Parsed data saved as Markdown: {md_file_path}")
        return {"success": True, "markdown_path": md_file_path}
    except Exception as e:
        logger.error(f"Failed to save parsed data as Markdown: {str(e)}")
        logger.error(traceback.format_exc())
        return {"success": False, "error": str(e)}

# 导入用于解析不同文件类型的依赖项
try:
    # 文档处理库
    from langchain_community.document_loaders import (
        PyPDFLoader, CSVLoader, Docx2txtLoader, UnstructuredExcelLoader,
        UnstructuredMarkdownLoader, UnstructuredPowerPointLoader,
        BSHTMLLoader, TextLoader, JSONLoader
    )
    from langchain_core.documents import Document
    has_langchain = True
except ImportError:
    logger.warning("未安装langchain相关库，某些高级解析功能可能不可用")
    has_langchain = False

try:
    # DOC文件处理
    from docx import Document as DocxDocument
    has_docx = True
except ImportError:
    logger.warning("未安装python-docx库，无法解析docx文件")
    has_docx = False

try:
    # PDF处理
    import pdfplumber
    has_pdfplumber = True
except ImportError:
    logger.warning("未安装pdfplumber库，PDF解析功能受限")
    has_pdfplumber = False

try:
    # PDF图像转换
    from pdf2image import convert_from_path
    has_pdf2image = True
except ImportError:
    logger.warning("未安装pdf2image库，无法从PDF提取图像")
    has_pdf2image = False

try:
    # OCR文本识别
    import pytesseract
    has_pytesseract = True
except ImportError:
    logger.warning("未安装pytesseract库，OCR功能不可用")
    has_pytesseract = False

try:
    # 图像处理
    from PIL import Image
    import cv2
    has_image_processing = True
except ImportError:
    logger.warning("未安装PIL或OpenCV库，图像处理功能受限")
    has_image_processing = False

try:
    # PPT处理
    from pptx import Presentation
    has_pptx = True
except ImportError:
    logger.warning("未安装python-pptx库，无法解析pptx文件")
    has_pptx = False

try:
    # HTML处理
    from bs4 import BeautifulSoup
    has_bs4 = True
except ImportError:
    logger.warning("未安装BeautifulSoup库，无法解析HTML文件")
    has_bs4 = False

try:
    # Markdown处理
    import markdown2
    has_markdown2 = True
except ImportError:
    logger.warning("未安装markdown2库，无法解析Markdown文件")
    has_markdown2 = False

try:
    # PDF高级处理
    import fitz  # PyMuPDF
    has_pymupdf = True
except ImportError:
    logger.warning("未安装PyMuPDF库，PDF高级解析功能不可用")
    has_pymupdf = False
try:
    # DOC转换
    import pypandoc
    has_pypandoc = True
except ImportError:
    logger.warning("未安装pypandoc库，无法进行doc转docx转换")
    has_pypandoc = False
# 检查 OCR 库
try:
    from paddleocr import PaddleOCR
    has_paddleocr = True
    global_paddle_ocr = PaddleOCR(use_angle_cls=True, lang="ch")
except:
    logger.warning("未安装PaddleOCR库，高级OCR功能不可用")
    has_paddleocr = False
    global_paddle_ocr = None

try:
    import easyocr
    has_easyocr = True
    easyocr_reader = easyocr.Reader(['ch_sim', 'en'])
except:
    logger.warning("未安装easy库，EasyOCR功能不可用")
    has_easyocr = False
    easyocr_reader = None

def allowed_file(filename):
    """检查文件是否具有允许的扩展名。如果文件名不包含扩展名，尝试猜测它的类型。"""
    if "." in filename:
        ext = filename.rsplit(".", 1)[1].lower()
        return ext in ALLOWED_EXTENSIONS
    else:
        # 文件名不包含扩展名时，使用 mimetypes 尝试猜测
        mime_type, _ = mimetypes.guess_type(filename)
        if mime_type:
            # 如果能猜测出 MIME 类型，检查是否在允许列表中
            if 'text/plain' in mime_type:
                return 'txt' in ALLOWED_EXTENSIONS
            elif 'image/png' in mime_type:
                return 'png' in ALLOWED_EXTENSIONS
            elif 'image/jpeg' in mime_type:
                return 'jpeg' in ALLOWED_EXTENSIONS
            elif 'application/vnd.ms-powerpoint' in mime_type:
                return 'ppt' in ALLOWED_EXTENSIONS
            elif 'application/msword' in mime_type:
                return 'doc' in ALLOWED_EXTENSIONS
        
        logger.warning(f"文件 {filename} 没有扩展名且无法猜测类型")
        return False  # 如果无法确定类型，默认不允许

# 解析 TXT 文件
def parse_txt(file_path):
    """解析文本文件，尝试多种编码"""
    try:
        encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1']
        text = None
        encoding_used = None
        
        # 尝试多种编码
        for encoding in encodings:
            try:
                with open(file_path, "r", encoding=encoding) as f:
                    text = f.read()
                    encoding_used = encoding
                    break
            except UnicodeDecodeError:
                continue
                
        if text is None:
            # 如果所有编码都失败，尝试以二进制模式读取
            with open(file_path, "rb") as f:
                text = f.read().decode('utf-8', errors='replace')
                encoding_used = 'utf-8 (with replacement)'
        
        return {
            "text": text,
            "encoding": encoding_used,
            "file_size_kb": os.path.getsize(file_path) / 1024
        }
    except Exception as e:
        logger.error(f"TXT 文件解析失败: {str(e)}")
        return {"error": f"TXT 文件解析失败: {str(e)}"}

# 直接从DOC文件中提取文本
def extract_text_from_doc(file_path):
    """尝试直接读取DOC文件内容"""
    try:
        # 尝试以二进制模式打开并寻找文本片段
        with open(file_path, 'rb') as f:
            content = f.read()
            
        # 尝试从二进制数据中提取文本
        text_chunks = []
        
        # 寻找ASCII文本片段
        i = 0
        while i < len(content) - 1:
            if (i + 20) < len(content):
                chunk = content[i:i+20]
                # 检查是否是可能的文本
                printable_chars = sum(1 for c in chunk if 32 <= c < 127)
                if printable_chars > 15:  # 如果大部分是可打印字符
                    # 找到文本块的起始和结束
                    start = i
                    while i < len(content) and (32 <= content[i] < 127 or content[i] in [9, 10, 13]):
                        i += 1
                    text_part = content[start:i].decode('ascii', errors='ignore')
                    if len(text_part.strip()) > 5:  # 忽略太短的片段
                        text_chunks.append(text_part.strip())
            i += 1
            
        # 查找Unicode文本
        i = 0
        while i < len(content) - 2:
            if content[i] == 0 and content[i+1] != 0 and i+2 < len(content) and content[i+2] == 0:
                text_start = i
                text_chars = []
                while i < len(content) - 2 and content[i] == 0 and content[i+1] != 0:
                    text_chars.append(content[i+1])
                    i += 2
                
                if len(text_chars) >= 5:
                    text = ''.join(chr(c) for c in text_chars if 32 <= c < 127)
                    if text.strip():
                        text_chunks.append(text)
            i += 1
            
        # 返回提取的文本
        if text_chunks:
            return {"text": "\n".join(text_chunks), "extraction_method": "直接二进制提取"}
        else:
            return {"text": "", "extraction_method": "直接二进制提取", "message": "未找到文本内容"}
    except Exception as e:
        logger.error(f"直接提取DOC文本失败: {str(e)}")
        return {"text": "", "extraction_method": "直接二进制提取失败", "error": str(e)}

# 使用 LibreOffice 将 DOC 转换为 DOCX
def convert_doc_to_docx(file_path):
    """使用 LibreOffice 将 DOC 转换为 DOCX"""
    if has_pypandoc:
        try:
            logger.info(f"使用 pypandoc 尝试转换 DOC: {file_path}")
            docx_path = str(file_path).replace(".doc", ".docx")
            pypandoc.convert_file(str(file_path), "docx", outputfile=docx_path)
            if os.path.exists(docx_path):
                logger.info(f"DOC 转换成功: {docx_path}")
                return {"success": True, "docx_path": docx_path}
            else:
                logger.warning(f"pypandoc 转换失败: {docx_path} 不存在")
        except Exception as e:
            logger.warning(f"pypandoc 转换失败: {str(e)}")
    
    # 如果 pypandoc 失败或不可用，尝试 LibreOffice
    try:
        # 可能的 LibreOffice 路径
        libreoffice_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            r"D:\LibreOffice\program\soffice.exe",
            "/usr/bin/soffice",
            "/usr/bin/libreoffice",
            "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        ]
        
        soffice_path = None
        for path in libreoffice_paths:
            if os.path.exists(path):
                soffice_path = path
                break
        
        if not soffice_path:
            logger.warning("未找到 LibreOffice 路径")
            return {"success": False, "error": "未找到 LibreOffice 路径"}
        
        logger.info(f"尝试使用 LibreOffice 转换 DOC: {file_path}")
        docx_path = str(file_path).replace(".doc", ".docx")
        abs_file_path = os.path.abspath(str(file_path))
        abs_output_dir = os.path.dirname(abs_file_path)
        
        cmd = [soffice_path, "--headless", "--convert-to", "docx", abs_file_path, "--outdir", abs_output_dir]
        logger.info(f"执行命令: {' '.join(cmd)}")
        
        process = subprocess.run(
            cmd, 
            capture_output=True, 
            text=True, 
            check=False,
            timeout=60
        )
        
        if os.path.exists(docx_path):
            logger.info(f"DOC 转换成功: {docx_path}")
            return {"success": True, "docx_path": docx_path}
        else:
            logger.warning(f"LibreOffice 转换失败: {docx_path} 不存在")
            return {"success": False, "error": f"转换失败: {process.stderr}"}
    
    except Exception as e:
        logger.error(f"DOC 转换失败: {str(e)}")
        return {"success": False, "error": str(e)}

# 解析 CSV 文件
def parse_csv(file_path):
    """解析 CSV 文件，尝试自动检测编码和分隔符"""
    try:
        # 尝试自动检测编码和分隔符
        encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1']
        separators = [',', '\t', ';', '|']
        
        for encoding in encodings:
            for sep in separators:
                try:
                    df = pd.read_csv(file_path, encoding=encoding, sep=sep)
                    if len(df.columns) > 1:  # 确保分隔符正确
                        # 转换数据为 Python 兼容格式
                        data = df.replace({np.nan: None}).to_dict(orient="records")
                        
                        return {
                            "data": data,
                            "info": {
                                "rows": len(df),
                                "columns": list(df.columns),
                                "encoding": encoding,
                                "separator": sep,
                                "dtypes": {str(col): str(dtype) for col, dtype in zip(df.columns, df.dtypes)}
                            }
                        }
                except Exception:
                    continue
                    
        # 如果所有尝试都失败，使用默认参数
        df = pd.read_csv(file_path)
        data = df.replace({np.nan: None}).to_dict(orient="records")
        
        return {
            "data": data,
            "info": {
                "rows": len(df),
                "columns": list(df.columns),
                "encoding": "自动检测",
                "separator": "自动检测",
                "dtypes": {str(col): str(dtype) for col, dtype in zip(df.columns, df.dtypes)}
            }
        }
    except Exception as e:
        logger.error(f"CSV 文件解析失败: {str(e)}")
        return {"error": f"CSV 文件解析失败: {str(e)}"}

# 解析 Markdown 文件
def parse_markdown(file_path):
    """解析 Markdown 文件，转换为 HTML 并保留原始文本"""
    if not has_markdown2:
        return {"error": "未安装 markdown2 库，无法解析 Markdown 文件"}
    
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            md_content = f.read()
        
        html_content = markdown2.markdown(md_content)
        
        # 提取标题和结构信息
        structure = []
        lines = md_content.split("\n")
        for line in lines:
            if line.startswith("#"):
                level = 0
                for char in line:
                    if char == '#':
                        level += 1
                    else:
                        break
                
                title = line[level:].strip()
                structure.append({"level": level, "title": title})
        
        return {
            "html": html_content,
            "text": md_content,
            "structure": structure
        }
    except Exception as e:
        logger.error(f"Markdown 文件解析失败: {str(e)}")
        return {"error": f"Markdown 文件解析失败: {str(e)}"}

# 解析 HTML 文件
def parse_html(file_path):
    """解析 HTML 文件，提取结构化内容"""
    if not has_bs4:
        return {"error": "未安装 BeautifulSoup 库，无法解析 HTML 文件"}
    
    try:
        with open(file_path, "r", encoding="utf-8", errors="replace") as f:
            html_content = f.read()
            soup = BeautifulSoup(html_content, "html.parser")
        
        # 提取标题
        title = soup.title.string if soup.title else ""
        
        # 提取所有文本
        text = soup.get_text(separator='\n', strip=True)
        
        # 提取链接
        links = []
        for link in soup.find_all('a'):
            href = link.get('href')
            if href:
                links.append({"text": link.text.strip(), "url": href})
                
        # 提取图片
        images = []
        for img in soup.find_all('img'):
            src = img.get('src')
            alt = img.get('alt', '')
            if src:
                images.append({"src": src, "alt": alt})
                
        # 提取表格
        tables = []
        for table in soup.find_all('table'):
            table_data = []
            for row in table.find_all('tr'):
                cells = row.find_all(['td', 'th'])
                if cells:
                    table_data.append([cell.text.strip() for cell in cells])
            if table_data:
                tables.append(table_data)
                
        # 提取结构信息
        headers = []
        for i in range(1, 7):
            for header in soup.find_all(f'h{i}'):
                headers.append({
                    "level": i,
                    "text": header.text.strip()
                })
        
        return {
            "title": title,
            "text": text,
            "links": links,
            "images": images,
            "tables": tables,
            "headers": headers,
            "html": html_content
        }
    except Exception as e:
        logger.error(f"HTML 文件解析失败: {str(e)}")
        return {"error": f"HTML 文件解析失败: {str(e)}"}

import os
from pdfplumber import open as pdfplumber_open  

def easyocr_image(image_path):
    
    tesseract_path = r'D:\Tesseract-ocr\tesseract.exe'
    pytesseract.pytesseract.tesseract_cmd = tesseract_path
    # 检查文件是否存在
    if not os.path.exists(image_path):
        return f"文件不存在，请检查文件路径: {image_path}"
    
    try:
        # 读取图像
        print("正在加载图像...")
        image = cv2.imread(image_path)
        
        # 如果图像无法被正确加载，返回错误信息
        if image is None:
            return "图像加载失败，请确认图像路径是否正确及图像文件是否损坏。"
        
        # 初始化 EasyOCR 识别器，支持中文和英文
        reader = easyocr.Reader(['ch_sim', 'en'])
        
        # 进行 OCR 识别
        result = reader.readtext(image)
        
        all_text = []
        
        # 提取文本块及其位置信息
        text_blocks = []
        for bbox, text, conf in result:  # 使用conf接收置信度
            if text.strip():  # 忽略空文本
                # 获取文本块的边界框坐标
                top_left = bbox[0]
                top_right = bbox[1]
                bottom_right = bbox[2]
                bottom_left = bbox[3]
                
                # 计算文本块的中心坐标和尺寸
                center_x = (top_left[0] + bottom_right[0]) / 2
                center_y = (top_left[1] + bottom_right[1]) / 2
                width = top_right[0] - top_left[0]
                
                text_blocks.append({
                    'text': text,
                    'center_x': center_x,
                    'center_y': center_y,
                    'top': min(top_left[1], top_right[1]),
                    'left': min(top_left[0], bottom_left[0]),
                    'width': width,
                    'bbox': bbox
                })
        
        # 使用更智能的行检测（基于聚类）
        # 按y坐标排序
        sorted_blocks = sorted(text_blocks, key=lambda x: x['top'])
        
        # 自适应行分组（基于文本块垂直重叠）
        rows = []
        current_row = [sorted_blocks[0]] if sorted_blocks else []
        
        for block in sorted_blocks[1:]:
            # 检查是否与当前行有足够的垂直重叠
            last_block = current_row[-1]
            last_bottom = last_block['bbox'][3][1]
            current_top = block['top']
            
            # 如果当前块的顶部位置低于上一个块的底部位置的一半，认为是新的一行
            overlap_threshold = (last_bottom - last_block['top']) * 0.5
            if current_top > last_bottom - overlap_threshold:
                # 完成当前行，开始新的一行
                rows.append(current_row)
                current_row = [block]
            else:
                # 添加到当前行
                current_row.append(block)
        
        # 添加最后一行
        if current_row:
            rows.append(current_row)
        
        # 处理每一行，保持原始的水平间距
        page_text = []
        
        # 获取页面宽度（用于规范化间距）
        page_width = image.shape[1]
        
        for row in rows:
            # 按x坐标排序行内的块
            sorted_row = sorted(row, key=lambda x: x['left'])
            
            # 创建带有空格的行文本，保持水平间距
            line_parts = []
            prev_end_x = 0
            
            for block in sorted_row:
                # 计算当前块与前一个块之间的间距
                start_x = block['left']
                if prev_end_x > 0:
                    # 计算空格数量，根据距离和平均字符宽度
                    char_width = 10  # 假设平均字符宽度为10像素
                    spaces = max(1, int((start_x - prev_end_x) / char_width))
                    line_parts.append(' ' * spaces)
                
                # 添加文本
                line_parts.append(block['text'])
                
                # 更新前一个块的结束位置
                prev_end_x = start_x + block['width']
            
            # 组合行文本
            row_text = ''.join(line_parts)
            
            # 处理缩进
            first_block = sorted_row[0] if sorted_row else None
            if first_block:
                indent_level = int(first_block['left'] / 50)  # 假设50像素为一个缩进级别
                indent = ' ' * (indent_level * 2)
                row_text = indent + row_text
            
            page_text.append(row_text)
        
        # 返回完整文本（所有页面）
        return "\n".join(page_text)
    except Exception as e:
        return f"发生错误: {e}"

def extract_text_with_positions(pdf_path):
    texts_by_page = []
    with fitz.open(pdf_path) as doc:
        for page in doc:
            blocks = page.get_text("blocks")
            text_items = []
            for block in blocks:
                x0, y0, x1, y1, text, *_ = block
                if text.strip():
                    text_items.append({
                        "type": "text",
                        "content": text.strip(),
                        "position": [round(y0, 2), round(x0, 2)]  # [top, left]
                    })
            texts_by_page.append(text_items)
    return texts_by_page

def extract_tables_with_positions(pdf_path):
    tables_by_page = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_tables = []
            try:
                tables = page.extract_tables()
                for table in tables:
                    if not table or not any(any(cell for cell in row) for row in table):
                        continue  # skip empty tables
                    bbox = page.bbox  # fallback if no position
                    top = bbox[1]
                    left = bbox[0]
                    page_tables.append({
                        "type": "table",
                        "content": table,
                        "position": [round(top, 2), round(left, 2)]
                    })
            except Exception:
                continue
            tables_by_page.append(page_tables)
    return tables_by_page

def extract_images_with_positions(pdf_path):
    images_by_page = []
    with fitz.open(pdf_path) as doc:
        for page_index, page in enumerate(doc):
            img_list = page.get_images(full=True)
            page_images = []
            for img_index, img in enumerate(img_list):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]
                bbox = page.get_image_bbox(img)
                top = bbox.y0
                left = bbox.x0
                page_images.append({
                    "type": "image",
                    "content": f"image_{page_index + 1}_{img_index + 1}.{image_ext}",
                    "image_bytes": image_bytes,
                    "position": [round(top, 2), round(left, 2)]
                })
            images_by_page.append(page_images)
    return images_by_page

def parse_pdf(file_path, max_pages=None, dpi=200):
    result = {
        "text": [],
        "tables": [],
        "images": [],
        "metadata": {},
        "content_by_position": []
    }

    try:
        logger.info(f"开始解析PDF文件: {file_path}")
        with pdfplumber_open(file_path) as pdf:
            total_pages = len(pdf.pages)
            result["metadata"] = {
                "pages": total_pages,
                "pdf_info": pdf.metadata
            }

            pdf_w, pdf_h = pdf.pages[0].width, pdf.pages[0].height
            print(f"pdf_h是：：{pdf_h}") # 添加输出

            for i, page in enumerate(pdf.pages):
                if max_pages and i >= max_pages:
                    break

                logger.info(f"处理第{i + 1}页")
                elements = []

                # 页面图像渲染
                images = convert_from_path(
                    file_path,
                    first_page=i + 1,
                    last_page=i + 1,
                    dpi=dpi,
                    poppler_path=r"D:\\poppler\\Library\\bin"
                )
                if not images:
                    continue
                page_image = images[0]
                img_w, img_h = page_image.size
                print(f"img_h是：：{img_h}") # 添加输出
                scale = img_w / pdf_w
                # 文本提取
                try:
                    text_blocks = page.extract_words()
                    current_y = None
                    text_content_list = []
                    for block in text_blocks:
                        if current_y is None or block['top'] != current_y:
                            current_y = block['top']
                            text_content_list.append({
                                'content': block['text'],
                                'y0': block['top']
                            })
                        else:
                            text_content_list[-1]['content'] += " " + block['text']

                    for text_block in text_content_list:
                        print(f"text的position是{str(text_block['y0'])}") # 添加输出
                        logger.info(f"text的position是{str(text_block['y0'])}") # 添加logger输出
                        elements.append({
                            "type": "text",
                            "content": text_block['content'],
                            "position": text_block['y0']
                        })
                    result["text"].append({
                        "page": i + 1,
                        "content": " ".join([block['content'] for block in text_content_list])
                    })
                except Exception as e:
                    result.setdefault("page_errors", {}).setdefault(i + 1, []).append(f"文本提取失败: {str(e)}")

                # 表格提取
                try:
                    page_tables = []
                    for table in page.find_tables():
                        table_data = table.extract()
                        if table_data:
                            page_tables.append(table_data)
                            top = table.bbox[3]# 修改
                            print(f"table的position是{str(top)}") # 添加输出
                            logger.info(f"table的position是{str(top)}") # 添加logger输出
                            elements.append({
                                "type": "table",
                                "content": table_data,
                                "position": top
                            })
                    if page_tables:
                        result["tables"].append({
                            "page": i + 1,
                            "tables": page_tables
                        })
                except Exception as e:
                    result.setdefault("page_errors", {}).setdefault(i + 1, []).append(f"表格提取失败: {str(e)}")

                # 图像提取与 OCR
                try:
                    for img_idx, img in enumerate(page.images):
                        x0, y0, x1, y1 = img["x0"], img["y0"], img["x1"], img["y1"]
                        left, right = x0 * scale, x1 * scale
                        top, bottom = (pdf_h - y1) * scale, (pdf_h - y0) * scale
                        cropped_image = page_image.crop((left, top, right, bottom))
                        temp_img_path = f"temp_img_{i}_{img_idx}.png"
                        cropped_image.save(temp_img_path)

                        extracted_text = ""
                        ocr_engine = "none"
                        print(f"Image coordinates: y0={img['y0']} y1={img['y1']} (PDF坐标)") # 提示
                        print(f"Using position: {img['y1']} for sorting") # 提示
                        
                        try:
                            if has_paddleocr and global_paddle_ocr:
                                ocr_result = global_paddle_ocr.ocr(temp_img_path, cls=True)
                                if ocr_result and ocr_result[0]:
                                    extracted_text = "\n".join(
                                        [line[1][0] for line in ocr_result[0] if line and line[1]]
                                    )
                                ocr_engine = "PaddleOCR"
                        finally:
                            if os.path.exists(temp_img_path):
                                os.remove(temp_img_path)

                        image_info = {
                            "x0": x0, "y0": y0, "x1": x1, "y1": y1,
                            "width": img["width"], "height": img["height"],
                            "page": i + 1, "index": img_idx,
                            "text": extracted_text, "ocr_engine": ocr_engine
                        }
                        print(f"image的position是{str(img['y1'])}") # 添加输出
                        logger.info(f"image的position是{str(img['y1'])}") # 添加logger输出
                        elements.append({
                            "type": "image",
                            "content": image_info,
                            "position": pdf_h-img["y1"]
                            #pdf_h pdf_h-img["y1"]修正position
                        })
                        result["images"].append(image_info)
                except Exception as e:
                    result.setdefault("page_errors", {}).setdefault(i + 1, []).append(f"图像提取失败: {str(e)}")

                # 排序后加入元素
                try:
                    elements.sort(key=lambda x: x['position'])
                    result["content_by_position"].append({
                        "page": i + 1,
                        "content": elements
                    })
                except Exception as e:
                    result.setdefault("page_errors", {}).setdefault(i + 1, []).append(f"内容排序失败: {str(e)}")

                # 内存调试
                mem_usage = psutil.Process(os.getpid()).memory_info().rss / 1024**2
                logger.info(f"[第{i + 1}页] 当前内存使用: {mem_usage:.2f} MB")

    except Exception as e:
        logger.error(f"使用 pdfplumber 解析 PDF 失败: {str(e)}")
        result["pdfplumber_error"] = str(e)

    return result

def extract_document_text(parsed_data, file_type=""):
    """
    从不同类型的文件解析结果中提取文本内容
    
    参数:
        parsed_data: 解析结果数据
        file_type: 文件类型
    
    返回:
        str: 提取的文本内容
    """
    text_content = ""
    
    if file_type == 'txt' and 'text' in parsed_data:
        # 文本文件直接获取文本内容
        text_content = parsed_data.get('text', '')
    
    elif file_type == 'pdf':
        # PDF文件可能有多种格式的文本内容
        if 'text' in parsed_data and isinstance(parsed_data['text'], list):
            # 格式1: 列表格式，每个元素包含页面内容
            text_content = ' '.join([page_text.get('content', '') for page_text in parsed_data['text'] if 'content' in page_text])
        elif 'content_by_position' in parsed_data and isinstance(parsed_data['content_by_position'], list):
            # 格式2: 按位置组织的内容
            all_texts = []
            for page in parsed_data['content_by_position']:
                for item in page.get('content', []):
                    if item['type'] == 'text' and isinstance(item.get('content', ''), str):
                        all_texts.append(item['content'])
            text_content = ' '.join(all_texts)
        # 如果以上都没有，尝试其他可能的文本字段
        elif 'text_content' in parsed_data:
            text_content = parsed_data['text_content']
    
    elif file_type == 'md' and 'text' in parsed_data:
        # Markdown文件
        text_content = parsed_data.get('text', '')
    
    elif file_type == 'html' and 'text' in parsed_data:
        # HTML文件
        text_content = parsed_data.get('text', '')
    
    elif file_type in ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff'] and 'text' in parsed_data:
        # 图像文件的OCR文本
        text_content = parsed_data.get('text', '')
    
    elif file_type == 'csv' and 'data' in parsed_data:
        # CSV文件尝试从数据构建文本
        data = parsed_data['data']
        if data and isinstance(data, list):
            # 将前几行数据转换为文本
            rows_text = []
            for row in data[:10]:  # 只取前10行
                if isinstance(row, dict):
                    row_text = ', '.join([f"{k}: {v}" for k, v in row.items()])
                    rows_text.append(row_text)
            text_content = '\n'.join(rows_text)
    
    elif file_type == 'json' and 'data' in parsed_data:
        # JSON文件尝试从数据构建文本
        data = parsed_data['data']
        if isinstance(data, dict):
            # 如果是对象，提取前几个键值对
            items = list(data.items())[:20]  # 只取前20个键值对
            text_content = '\n'.join([f"{k}: {v}" for k, v in items])
        elif isinstance(data, list) and data:
            # 如果是数组，提取前几个元素的信息
            if isinstance(data[0], dict):
                items = list(data[0].items())[:20]  # 只取第一个对象的前20个键值对
                text_content = '\n'.join([f"{k}: {v}" for k, v in items])
            else:
                # 如果数组元素不是对象，直接转换前几个元素
                text_content = '\n'.join([str(item) for item in data[:20]])
    
    return text_content

def summarize_with_ollama(text, file_info=None, use_ollama=True):
    """
    使用Ollama API生成文本内容摘要，如果不可用则使用本地方法
    
    参数:
        text: 需要摘要的文本内容
        file_info: 文件信息
        use_ollama: 是否尝试使用Ollama API
    
    返回:
        str: 生成的摘要
    """
    if not text or len(text.strip()) < 50:
        return "文档内容过短或无法提取有效内容进行摘要。"
    
    filename = file_info.get('filename', '未知文件') if file_info else '未知文件'
    file_type = file_info.get('file_type', '') if file_info else ''
    
    # 构建提示
    prompt = f"""
    请对以下文件内容生成一个详细、准确的摘要，包括文档的主要要点和关键信息：
    文件名：{filename}
    文件类型：{file_type}
    内容：
    {text[:5000]}  # 限制内容长度，防止请求过大
    """
    if use_ollama:
        try:
            ollama_url = 'http://10.2.60.75:11436/api/generate'  
            data = {
                'model': 'deepseek-r1:1.5b',  
                "prompt": prompt,
                "stream": False
            }        
            response = requests.post(ollama_url, json=data, timeout=10)
            if response.status_code == 200:
                return response.json().get('response', "")
        except Exception as e:
            # 如果API调用失败，记录错误并回退到本地方法
            print(f"Ollama API调用失败: {str(e)}，使用本地摘要方法")
    text = text.replace('\n', ' ').replace('\r', ' ')
    while '  ' in text:
        text = text.replace('  ', ' ')
    text = text.strip()
    sentences = []
    for s in text.split('. '):
        s = s.strip()
        if s:
            if not s.endswith('.'):
                s += '.'
            sentences.append(s)
    
    if not sentences:
        return "无法从文本中提取有效的句子进行摘要。"
    summary_sentences = []
    
    # 包含第一句
    if sentences:
        summary_sentences.append(sentences[0])
    if len(sentences) <= 5:
        summary_sentences = sentences
    else:
        # 从文档的不同部分选择句子
        middle_idx = len(sentences) // 2
        if middle_idx < len(sentences) and sentences[middle_idx] not in summary_sentences:
            summary_sentences.append(sentences[middle_idx])
        
        # 从末尾部分选择
        if len(sentences) >= 3 and sentences[-2] not in summary_sentences:
            summary_sentences.append(sentences[-2])
            
        candidates = [(i, len(s.split())) for i, s in enumerate(sentences) if s not in summary_sentences]
        candidates.sort(key=lambda x: x[1], reverse=True)  # 按句子长度排序
        
        # 添加前几个最长的句子，直到有5个句子
        for i, _ in candidates[:5-len(summary_sentences)]:
            summary_sentences.append(sentences[i])
    
    # 按原始顺序重新排序句子
    summary_sentences = [s for s in sentences if s in summary_sentences][:5]  # 限制最多5个句子
    
    # 组合成摘要
    summary = ' '.join(summary_sentences)
    
    # 格式化返回
    return summary

def generate_content_summary(parsed_data, file_info=None):
    """
    生成文件内容的文本摘要
    
    参数:
        parsed_data: 解析结果数据
        file_info: 文件基本信息
    
    返回:
        str: 生成的内容摘要Markdown文本
    """
    # 获取文件类型
    file_type = file_info.get('file_type', '').lower() if file_info else ''
    
    # 提取文本内容
    text_content = extract_document_text(parsed_data, file_type)
    
    if not text_content or len(text_content.strip()) < 50:
        return "## 内容摘要\n\n无法从文件中提取足够的文本内容进行摘要。\n"
    
    # 生成摘要
    try:
        summary = summarize_with_ollama(text_content, file_info, use_ollama=False) # 设置为False以禁用API调用
        return f"## 内容摘要\n\n{summary}\n"
    except Exception as e:
        return f"## 内容摘要\n\n生成摘要时出错: {str(e)}\n"

# 使用 pytesseract 从图像中提取文本
def extract_text_from_image(image_path):
    """使用 pytesseract 从图像中提取文本"""
    if not has_pytesseract or not has_image_processing:
        return ""
    
    try:
        # 读取图像
        img = cv2.imread(str(image_path))
        
        # 如果OpenCV读取失败，尝试使用PIL
        if img is None:
            pil_img = Image.open(str(image_path))
            img = np.array(pil_img)
        
        # 转为灰度
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # 应用阈值处理
        gray = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
        
        # OCR
        text = pytesseract.image_to_string(gray, lang="eng+chi_sim")
        return text
    except Exception as e:
        logger.error(f"图像OCR提取失败: {str(e)}")
        return ""

# 使用 PaddleOCR 从图像中提取文本
def extract_text_with_paddleocr(image_path):
    """使用 PaddleOCR 从图像中提取文本"""
    if not has_paddleocr:
        return ""
    
    try:
        ocr = PaddleOCR(use_angle_cls=True, lang="ch")  # 支持中文识别
        result = ocr.ocr(str(image_path), cls=True)
        
        text_lines = []
        
        if result and len(result) > 0:
            for idx, res in enumerate(result):
                if res:
                    page_text = []
                    
                    for line in res:
                        text = line[1][0]  # 识别的文本
                        page_text.append(text)
                    
                    text_lines.append("\n".join(page_text))
        
        return "\n".join(text_lines)
    except Exception as e:
        logger.error(f"PaddleOCR提取失败: {str(e)}")
        return ""

# 解析图像文件
def parse_image(file_path):
    """解析图像文件，提取文本和元数据"""
    if not has_image_processing:
        return {"error": "未安装图像处理库"}
    
    try:
        # 获取图像基本信息
        img = Image.open(file_path)
        width, height = img.size
        format_type = img.format
        mode = img.mode
        
        result = {
            "image_info": {
                "width": width,
                "height": height,
                "format": format_type,
                "mode": mode,
                "file_size_kb": os.path.getsize(file_path) / 1024,
            }
        }
        
        # 提取EXIF数据（如果有）
        if hasattr(img, '_getexif') and img._getexif():
            exif = img._getexif()
            if exif:
                result["exif"] = {str(k): str(v) for k, v in exif.items()}
        
        # 使用OCR提取文本
        if has_paddleocr:
            result["text"] = extract_text_with_paddleocr(file_path)
            result["ocr_engine"] = "PaddleOCR"
        elif has_pytesseract:
            result["text"] = extract_text_from_image(file_path)
            result["ocr_engine"] = "Tesseract"
        else:
            result["text"] = ""
            result["ocr_engine"] = "未安装OCR引擎"
        
        return result
    except Exception as e:
        logger.error(f"图像解析失败: {str(e)}")
        return {"error": f"图像解析失败: {str(e)}"}

# 解析 JSON 文件
def parse_json(file_path):
    """解析 JSON 文件"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # 分析JSON结构
        if isinstance(data, dict):
            structure = "对象"
            keys = list(data.keys())
            if len(keys) <= 20:
                top_level_keys = keys
            else:
                top_level_keys = keys[:20] + [f"...其余 {len(keys) - 20} 个键"]
        elif isinstance(data, list):
            structure = "数组"
            top_level_keys = []
            length = len(data)
            if length > 0 and isinstance(data[0], dict) and len(data) <= 1000:
                # 分析数组中对象的键
                all_keys = set()
                for item in data[:100]:  # 只分析前100个项目
                    if isinstance(item, dict):
                        all_keys.update(item.keys())
                top_level_keys = list(all_keys)
                if len(top_level_keys) > 20:
                    top_level_keys = top_level_keys[:20] + [f"...其余 {len(top_level_keys) - 20} 个键"]
        else:
            structure = "基本类型"
            top_level_keys = []
        
        return {
            "structure": structure,
            "top_level_keys": top_level_keys,
            "size_kb": os.path.getsize(file_path) / 1024,
            "data": data
        }
    except Exception as e:
        logger.error(f"JSON 文件解析失败: {str(e)}")
        return {"error": f"JSON 文件解析失败: {str(e)}"}

# 解析上传文件的统一接口
def parse_file(file_path: str, original_filename=None, userId=None):
    """根据文件类型解析文件内容"""
    start_time = time.time()
    
    # 使用文件名（不包含路径）获取扩展名
    filename = os.path.basename(file_path)
    logger.info(f"正在解析文件: {filename}, 路径: {file_path}")

    if original_filename:
        logger.info(f"原始文件名: {original_filename}")
        original_name_without_ext = os.path.splitext(original_filename)[0]
        display_name = original_name_without_ext
    # else:      
    #     display_name = os.path.splitext(filename)[0]     # 如果没传 original_filename，就退回到安全文件名去掉扩展名

    if not os.path.exists(file_path):
        return {"error": f"文件不存在: {file_path}"}
        
    # 尝试从文件路径确定扩展名
    if "." in filename:
        ext = filename.rsplit(".", 1)[-1].lower()
    else:      # 如果文件名中没有扩展名，尝试通过MIME类型识别
        try:
            mime_type, _ = mimetypes.guess_type(filename)
            if mime_type:
                if 'text/plain' in mime_type:
                    ext = 'txt'
                elif 'image/png' in mime_type:
                    ext = 'png'
                elif 'image/jpeg' in mime_type:
                    ext = 'jpeg'
                elif 'application/vnd.ms-powerpoint' in mime_type:
                    ext = 'ppt'
                elif 'application/msword' in mime_type:
                    ext = 'doc'
                else:
                    ext = 'unknown'
            else:
                ext = 'unknown'
        except Exception:
            ext = 'unknown'
    
    logger.info(f"识别的文件扩展名: {ext}")

    # 提取文件基本信息
    file_info = {
        "filename": original_filename if original_filename else filename,  # Use the FULL original filename
        "original_filename": original_filename if original_filename else filename,
        "file_size_kb": os.path.getsize(file_path) / 1024,
        "file_type": ext,
        "last_modified": os.path.getmtime(file_path)
    }

    # 根据文件类型解析内容
    result = {"file_info": file_info}
    pdf_path=file_path
    try:
        # 对于非图像和非PDF文件，尝试转换为PDF
        pdf_result = None
        if ext not in ["png", "jpg", "jpeg", "gif", "bmp", "tiff", "pdf", "csv", "json", "txt", "html", "md"]:
            logger.info(f"开始将文件转换为PDF: {file_path}")
            pdf_result = convert_to_pdf(file_path, original_filename)
            logger.info(f"PDF转换结果: {pdf_result}")
            
            # 如果PDF转换成功，解析PDF
            if pdf_result.get("success", False):
                pdf_path = pdf_result["pdf_path"]
                logger.info(f"开始解析转换后的PDF: {pdf_path}")
                result.update(parse_pdf(pdf_path))
                logger.info(f"PDF解析完成")
            else:
                logger.warning(f"PDF转换失败: {pdf_result.get('error', '未知错误')}")
                result["pdf_conversion"] = {
                    "success": False,
                    "error": pdf_result.get("error", "PDF转换失败")
                }
        elif ext == "pdf":
            result.update(parse_pdf(file_path))
        elif ext == "md":
            result.update(parse_markdown(file_path))
        elif ext == "html":
            result.update(parse_html(file_path))
        elif ext in ["png", "jpg", "jpeg", "gif", "bmp", "tiff"]:
            result.update(parse_image(file_path))
        elif ext == "json":
            result.update(parse_json(file_path))
        elif ext == "txt":
            result.update(parse_txt(file_path))
        elif ext == "csv":
            result.update(parse_csv(file_path))
        else:
            # 如果没有支持的解析器但PDF转换成功，使用PDF数据作为结果
            if pdf_result and pdf_result.get("success", False):
                logger.info(f"对于不支持的文件类型使用PDF转换数据: {ext}")
            else:
                result["error"] = f"不支持的文件格式: {ext}"

    except Exception as e:
        logger.error(f"解析文件失败: {str(e)}")
        logger.error(traceback.format_exc())
        result["error"] = f"解析文件失败: {str(e)}"
    
    # 添加处理时间
    result["processing_time"] = time.time() - start_time

    original_name_without_ext = os.path.splitext(original_filename)[0]
    md_result = save_parsed_data_as_markdown(result, pdf_path, original_name_without_ext, userId)
    logger.info(f"Markdown保存结果: {md_result}")
    return result

# 上传文件接口
@app.route("/upload", methods=["POST"])
def upload_file():
    """Upload and parse file"""
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file part"}), 400

        file = request.files["file"]
        if file.filename == "":
            return jsonify({"error": "No file selected"}), 400
        format_type = request.args.get('format', 'json')
        userId = request.args.get('userId')
        if file and allowed_file(file.filename):
            original_filename = file.filename
            original_extension = original_filename.rsplit(".", 1)[-1].lower() if "." in original_filename else ""

            safe_filename = os.path.splitext(original_filename)[0]
            if "." not in safe_filename and original_extension:
                safe_filename = f"{safe_filename}.{original_extension}"
                
            logger.info(f"Original filename: {original_filename}, Safe filename: {safe_filename}")
            
            file_path = os.path.join(app.config["UPLOAD_FOLDER"], safe_filename)
            file.save(file_path)
            
            # Verify file was saved correctly
            if not os.path.exists(file_path):
                return jsonify({"error": "File save failed"}), 500
            parsed_data = parse_file(file_path, original_filename, userId)

            #将文件类型替换为pdf，否则convert_to_markdown会报错（与file.type不符）
            if original_extension in ["docx", "xlsx", "xls", "doc", "ppt", "pptx"]:
                safe_filename_record=file.filename
                extension_record="pdf"
                safe_filename_find = f"{safe_filename_record}.{extension_record}"
            
            #file_path_record=os.path.join(app.config["UPLOAD_FOLDER"], safe_filename_find)
                logger.info(f"safe_filename_recordline1613!!: {safe_filename_record}, file.filename: {file.filename}")
                logger.info(f"extension_record: {extension_record}, extension_record: {extension_record}")
                logger.info(f"safe_filename_find: {safe_filename_find}, safe_filename_find: {safe_filename_find}")
              # 文件信息
                parsed_data["file_info"]["file_type"] = extension_record

            # Ensure file_info contains both the original and safe filename
            parsed_data["file_info"]["original_filename"] = original_filename
            parsed_data["file_info"]["safe_filename"] = safe_filename
            parsed_data["file_info"]["filename"] = original_filename  # Use original for display

            
            # Return according to requested format
            if format_type.lower() == 'markdown':
                # Generate Markdown output
                markdown_content = convert_to_markdown(parsed_data, parsed_data["file_info"])
                return Response(markdown_content, mimetype='text/markdown')
            else:
                # Default to JSON
                return jsonify({
                    "message": "File uploaded and parsed successfully", 
                    "filename": original_filename,
                    "safe_filename": safe_filename,
                    "parsed_data": parsed_data
                })

        return jsonify({"error": "Unsupported file format"}), 400
    
    except Exception as e:
        logger.error(f"File upload processing failed: {str(e)}")
        logger.error(traceback.format_exc())
        return jsonify({"error": f"File upload processing failed: {str(e)}"}), 500

# 获取已上传文件列表接口
@app.route("/files", methods=["GET"])
def list_files():
    """列出所有上传的文件"""
    try:
        files = []
        for filename in os.listdir(app.config["UPLOAD_FOLDER"]):
            file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            if os.path.isfile(file_path):
                file_size = os.path.getsize(file_path) / 1024  # KB
                last_modified = os.path.getmtime(file_path)
                ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
                
                files.append({
                    "filename": filename,
                    "file_size_kb": file_size,
                    "file_type": ext,
                    "last_modified": last_modified
                })
                
        return jsonify({"files": files})
    except Exception as e:
        logger.error(f"列出文件失败: {str(e)}")
        return jsonify({"error": f"列出文件失败: {str(e)}"}), 500

# 获取指定文件解析内容接口
@app.route("/files/<path:filename>", methods=["GET"])
def get_file_content(filename):
    """获取特定文件的解析内容"""
    try:
        # 检查是否请求Markdown格式
        format_type = request.args.get('format', 'json')
        
        # 使用 path:filename 处理可能包含特殊字符的文件名
        file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        logger.info(f"请求获取文件内容: {filename}, 完整路径: {file_path}")
        
        if not os.path.exists(file_path):
            logger.error(f"文件不存在: {file_path}")
            return jsonify({"error": "文件不存在"}), 404
        
        if not os.path.isfile(file_path):
            logger.error(f"路径不是文件: {file_path}")
            return jsonify({"error": "路径不是文件"}), 400

        # 尝试获取原始文件名（如果有的话）
        original_filename = None
        json_path = os.path.splitext(file_path)[0] + ".json"
        if os.path.exists(json_path):
            try:
                with open(json_path, 'r', encoding='utf-8') as f:
                    json_data = json.load(f)
                    if "file_info" in json_data and "original_filename" in json_data["file_info"]:
                        original_filename = json_data["file_info"]["original_filename"]
            except Exception as e:
                logger.warning(f"读取JSON文件获取原始文件名失败: {str(e)}")

        # 解析文件内容，传入原始文件名
        parsed_data = parse_file(file_path, original_filename)
        
        # 文件信息
        file_info = {
            "filename": filename,
            "original_filename": original_filename if original_filename else filename,
            "file_type": filename.rsplit(".", 1)[-1].lower() if "." in filename else "",
            "file_size_kb": os.path.getsize(file_path) / 1024,
            "last_modified": os.path.getmtime(file_path)
        }
        
        # 根据请求的格式返回
        if format_type.lower() == 'markdown':
            # 生成Markdown输出
            markdown_content = convert_to_markdown(parsed_data, file_info)
            return Response(markdown_content, mimetype='text/markdown')
        else:
            # 默认返回JSON
            return jsonify({"filename": filename, "original_filename": original_filename, "parsed_data": parsed_data})
    except Exception as e:
        logger.error(f"获取文件内容失败: {str(e)}")
        return jsonify({"error": f"获取文件内容失败: {str(e)}"}), 500

# 删除指定文件接口
@app.route("/files/<path:filename>", methods=["DELETE"])
def delete_file(filename):
    """删除特定文件"""
    try:
        file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        if not os.path.exists(file_path):
            return jsonify({"error": "文件不存在"}), 404
            
        os.remove(file_path)
        return jsonify({"message": f"文件 {filename} 已成功删除"})
    except Exception as e:
        logger.error(f"删除文件失败: {str(e)}")
        return jsonify({"error": f"删除文件失败: {str(e)}"}), 500

def escape_md_table_cell(text):
    if not text:
        return ""
    escaped = str(text).replace("|", "\\|")
    escaped = escaped.replace("\n", " ")
    if len(escaped) > 100:
        escaped = escaped[:97] + "..."
        
    return escaped

# Markdown转换函数
def convert_to_markdown(parsed_data, file_info=None, full_content=True):
    """
    将解析的数据转换为Markdown格式
    
    参数:
        parsed_data: 解析结果数据
        file_info: 文件基本信息
        full_content: 是否输出完整内容（而非截断版本）
    """
    md_output = []

    # 添加标题和文件信息
    if file_info:
        md_output.append(f"# 文件分析报告: {file_info.get('filename', '未知文件')}")
        md_output.append("")
        md_output.append("## 文件信息")
        md_output.append("")
        md_output.append(f"- **文件名**: {file_info.get('filename', '未知')}")
        md_output.append(f"- **文件类型**: {file_info.get('file_type', '未知')}")
        md_output.append(f"- **文件大小**: {file_info.get('file_size_kb', 0):.2f} KB")
        
        # 格式化最后修改时间
        if 'last_modified' in file_info:
            last_modified = datetime.fromtimestamp(file_info['last_modified']).strftime('%Y-%m-%d %H:%M:%S')
            md_output.append(f"- **最后修改**: {last_modified}")
        
        md_output.append("")
    
    # 添加内容总结（在文件信息之后，详细内容之前） 
    content_summary = generate_content_summary(parsed_data, file_info)
    md_output.append(content_summary)

    # 处理可能的错误
    if "error" in parsed_data:
        md_output.append("## 错误")
        md_output.append("")
        md_output.append(f"```\n{parsed_data['error']}\n```")
        md_output.append("")
        return "\n".join(md_output)
    
    # 根据文件类型构建不同的Markdown内容
    file_type = file_info.get('file_type', '').lower() if file_info else ''
    
    # 处理文本文件
    if file_type == 'txt' and 'text' in parsed_data:
        md_output.append("## 文本内容")
        md_output.append("")
        md_output.append("```")
        md_output.append(parsed_data.get('text', ''))
        md_output.append("```")
        
        if 'encoding' in parsed_data:
            md_output.append(f"\n*文件编码: {parsed_data['encoding']}*")
    
    # 处理CSV文件
    elif file_type == 'csv' and 'data' in parsed_data:
        md_output.append("## CSV内容")
        md_output.append("")
        
        if 'info' in parsed_data:
            info = parsed_data['info']
            md_output.append(f"- **行数**: {info.get('rows', 0)}")
            md_output.append(f"- **列数**: {len(info.get('columns', []))}")
            if 'encoding' in info:
                md_output.append(f"- **编码**: {info['encoding']}")
            if 'separator' in info:
                separator_display = info['separator'].replace('\t', '\\t')
                md_output.append(f"- **分隔符**: '{separator_display}'")
            md_output.append("")
        
        # 转换数据为Markdown表格
        data = parsed_data['data']
        if data and len(data) > 0:
            # 获取列名
            columns = parsed_data.get('info', {}).get('columns', list(data[0].keys()))
            
            # 创建表头
            md_output.append("| " + " | ".join(str(col) for col in columns) + " |")
            md_output.append("| " + " | ".join(["---"] * len(columns)) + " |")
            
            # 添加数据行
            row_limit = len(data) if full_content else min(15, len(data))
            for row in data[:row_limit]:
                row_values = []
                for col in columns:
                    val = str(row.get(col, "")).replace("|", "\\|").replace("\n", " ")
                    # 截断过长的单元格内容
                    if len(val) > 80 and not full_content:
                        val = val[:77] + "..."
                    row_values.append(val)
                md_output.append("| " + " | ".join(row_values) + " |")
            
            # 如果有更多行且不是完整内容模式
            if len(data) > row_limit and not full_content:
                md_output.append("\n*显示前" + str(row_limit) + "行数据，共 " + str(len(data)) + " 行*")
    
    # 处理PDF文件
    elif file_type == 'pdf':
        md_output.append("## PDF内容分析")
        md_output.append("")
        
        # 添加元数据
        if 'metadata' in parsed_data:
            md_output.append("### 文档信息")
            md_output.append("")
            
            metadata = parsed_data['metadata']
            if 'pages' in metadata:
                md_output.append(f"- **页数**: {metadata['pages']}")
            
            pdf_info = metadata.get('pdf_info', {})
            for key, value in pdf_info.items():
                if value:
                    md_output.append(f"- **{key}**: {value}")
            
            pymupdf_info = metadata.get('pymupdf_info', {})
            if pymupdf_info:
                md_output.append("")
                md_output.append("#### PyMuPDF元数据")
                for key, value in pymupdf_info.items():
                    if value:
                        md_output.append(f"- **{key}**: {value}")
            
            md_output.append("")
        
        # 使用按位置组织的内容
        if 'content_by_position' in parsed_data and parsed_data['content_by_position']:
            md_output.append("### 文档内容（按原始位置排序）")
            md_output.append("")
            
            page_limit = len(parsed_data['content_by_position']) if full_content else min(5, len(parsed_data['content_by_position']))
            
            for page in parsed_data['content_by_position'][:page_limit]:
                md_output.append(f"#### 第 {page.get('page', '')} 页")
                md_output.append("")
                
                for item in page.get('content', []):
                    if item['type'] == 'text':
                        content = item.get('content', '')
                        if len(content) > 1500 and not full_content:
                            content = content[:1500] + "\n... (内容已截断)"
                        md_output.append(content)
                        md_output.append("")
                    elif item['type'] == 'image':
                        img_info = item.get('content', {})
                        md_output.append("*[图像]*")
                        md_output.append(f"- 尺寸: {img_info.get('width', 'N/A')} x {img_info.get('height', 'N/A')}")
                        if 'text' in img_info and img_info['text']:
                            md_output.append(f"- OCR文本: {img_info['text']}")
                        md_output.append("")
                    elif item['type'] == 'table':
                        md_output.append("*[表格]*")
                        md_output.append(convert_table_to_markdown(item.get('content', []), full_content))
                        md_output.append("")
                
                md_output.append("")
    
        
    # 处理Markdown文件
    elif file_type == 'md' and ('text' in parsed_data or 'html' in parsed_data):
        md_output.append("## Markdown内容")
        md_output.append("")
        
        if 'text' in parsed_data:
            md_output.append("```markdown")
            md_output.append(parsed_data.get('text', ''))
            md_output.append("```")
            md_output.append("")
        
        # 添加HTML渲染结果
        if full_content and 'html' in parsed_data:
            md_output.append("### HTML渲染结果")
            md_output.append("")
            md_output.append("```html")
            html_content = parsed_data.get('html', '')
            md_output.append(html_content)
            md_output.append("```")
            md_output.append("")
        
        # 添加结构信息
        if 'structure' in parsed_data and parsed_data['structure']:
            md_output.append("### 文档结构")
            md_output.append("")
            for item in parsed_data['structure']:
                level = item.get('level', 1)
                title = item.get('title', '')
                md_output.append("  " * (level-1) + f"- {title}")
            md_output.append("")
    
    # 处理HTML文件
    elif file_type == 'html':
        md_output.append("## HTML内容分析")
        md_output.append("")
        
        # 添加标题
        if 'title' in parsed_data and parsed_data['title']:
            md_output.append(f"**页面标题**: {parsed_data['title']}")
            md_output.append("")
        
        # 添加文本预览
        if 'text' in parsed_data and parsed_data['text']:
            md_output.append("### 文本内容")
            md_output.append("")
            text = parsed_data['text']
            preview_len = len(text) if full_content else min(1500, len(text))
            preview_text = text[:preview_len]
            md_output.append("```")
            md_output.append(preview_text)
            if len(text) > preview_len and not full_content:
                md_output.append("... (内容已截断)")
            md_output.append("```")
            md_output.append("")
        
        # 添加标题结构
        if 'headers' in parsed_data and parsed_data['headers']:
            md_output.append("### 页面结构")
            md_output.append("")
            
            for header in parsed_data['headers']:
                level = header.get('level', 1)
                text = header.get('text', '')
                md_output.append("  " * (level-1) + f"- {text}")
            
            md_output.append("")
        
        # 添加链接信息
        if 'links' in parsed_data and parsed_data['links']:
            md_output.append("### 链接")
            md_output.append("")
            
            link_limit = len(parsed_data['links']) if full_content else min(20, len(parsed_data['links']))
            
            for i, link in enumerate(parsed_data['links'][:link_limit]):
                text = link.get('text', '')
                url = link.get('url', '')
                md_output.append(f"{i+1}. [{text}]({url})")
            
            # 如果有更多链接且不是完整内容模式
            if len(parsed_data['links']) > link_limit and not full_content:
                md_output.append(f"\n*只显示前{link_limit}个链接，共 {len(parsed_data['links'])} 个*")
            
            md_output.append("")
        
        # 添加图片信息
        if 'images' in parsed_data and parsed_data['images']:
            md_output.append("### 图片")
            md_output.append("")
            md_output.append(f"共 {len(parsed_data['images'])} 张图片")
            md_output.append("")
            
            image_limit = len(parsed_data['images']) if full_content else min(10, len(parsed_data['images']))
            
            for i, img in enumerate(parsed_data['images'][:image_limit]):
                src = img.get('src', '')
                alt = img.get('alt', '')
                md_output.append(f"{i+1}. ![{alt}]({src})")
            
            # 如果有更多图片且不是完整内容模式
            if len(parsed_data['images']) > image_limit and not full_content:
                md_output.append(f"\n*只显示前{image_limit}张图片，共 {len(parsed_data['images'])} 张*")
            
            md_output.append("")
        
        # 添加表格信息
        if 'tables' in parsed_data and parsed_data['tables']:
            md_output.append("### 表格")
            md_output.append("")
            md_output.append(f"共 {len(parsed_data['tables'])} 个表格")
            md_output.append("")
            
            table_limit = len(parsed_data['tables']) if full_content else min(5, len(parsed_data['tables']))
            
            for i, table in enumerate(parsed_data['tables'][:table_limit]):
                md_output.append(f"**表格 {i+1}**:")
                md_output.append("")
                md_output.append(convert_table_to_markdown(table, full_content))
                md_output.append("")
            
            # 如果有更多表格且不是完整内容模式
            if len(parsed_data['tables']) > table_limit and not full_content:
                md_output.append(f"*只显示前{table_limit}个表格*")
                md_output.append("")
        
        # 如果展示完整内容，添加完整HTML源码
        if full_content and 'html' in parsed_data:
            md_output.append("### HTML源码")
            md_output.append("")
            md_output.append("```html")
            md_output.append(parsed_data['html'])
            md_output.append("```")
            md_output.append("")
    
    # 处理图像文件
    elif file_type in ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff'] and 'image_info' in parsed_data:
        md_output.append("## 图像分析")
        md_output.append("")
        
        # 图像基本信息
        image_info = parsed_data.get('image_info', {})
        md_output.append("### 图像信息")
        md_output.append("")
        md_output.append(f"- **尺寸**: {image_info.get('width', 0)} x {image_info.get('height', 0)}")
        md_output.append(f"- **格式**: {image_info.get('format', '未知')}")
        md_output.append(f"- **模式**: {image_info.get('mode', '未知')}")
        md_output.append("")
        
        # OCR识别文本
        if 'text' in parsed_data and parsed_data['text']:
            md_output.append("### OCR识别文本")
            md_output.append("")
            md_output.append("```")
            md_output.append(parsed_data['text'])
            md_output.append("```")
            md_output.append("")
            
            if 'ocr_engine' in parsed_data:
                md_output.append(f"*使用 {parsed_data['ocr_engine']} 引擎识别*")
                md_output.append("")
        
        # EXIF数据
        if 'exif' in parsed_data and parsed_data['exif']:
            md_output.append("### EXIF元数据")
            md_output.append("")
            
            for key, value in parsed_data['exif'].items():
                md_output.append(f"- **{key}**: {value}")
            md_output.append("")
    
    # 处理JSON文件
    elif file_type == 'json' and 'data' in parsed_data:
        md_output.append("## JSON内容分析")
        md_output.append("")
        
        # 结构信息
        if 'structure' in parsed_data:
            md_output.append(f"- **结构类型**: {parsed_data['structure']}")
        if 'top_level_keys' in parsed_data:
            md_output.append(f"- **顶级键**: {', '.join(str(key) for key in parsed_data['top_level_keys'])}")
        if 'size_kb' in parsed_data:
            md_output.append(f"- **文件大小**: {parsed_data['size_kb']:.2f} KB")
        md_output.append("")
        
        # JSON内容
        md_output.append("### JSON内容")
        md_output.append("")
        md_output.append("```json")
        
        # 处理JSON数据
        try:
            # 漂亮地格式化JSON以便于阅读
            data_str = json.dumps(parsed_data['data'], indent=2, ensure_ascii=False)
            
            # 如果不是完整内容模式且内容很长，则截断
            if not full_content and len(data_str) > 5000:
                data_str = data_str[:5000] + "\n... (内容已截断)"
            
            md_output.append(data_str)
        except Exception as e:
            # 如果JSON格式化失败，使用字符串表示
            data_str = str(parsed_data['data'])
            if not full_content and len(data_str) > 5000:
                data_str = data_str[:5000] + "... (内容已截断)"
            md_output.append(data_str)
            md_output.append(f"\n*JSON格式化错误: {str(e)}*")
        
        md_output.append("```")
        md_output.append("")
    
    # 捕获所有剩余的解析字段，确保不丢失信息
    remaining_fields = [field for field in parsed_data.keys() if field not in [
        'file_info', 'error', 'text', 'metadata', 'tables', 'images_info', 'sheets', 'data', 'info',
        'html', 'structure', 'slides', 'properties', 'images', 'links', 'headers', 'title',
        'image_info', 'exif', 'ocr_engine', 'structure', 'top_level_keys', 'size_kb',
        'pymupdf_images', 'ocr_text', 'text_source', 'conversion_method', 'processing_time'
    ]]
    
    # 添加处理时间
    if 'processing_time' in parsed_data:
        md_output.append("---")
        md_output.append(f"*处理时间: {parsed_data['processing_time']:.2f} 秒*")
    
    return "\n".join(md_output)

def convert_table_to_markdown(table, full_content=False):
    """
    将表格数据转换为Markdown表格格式
    
    参数:
        table: 表格数据
        full_content: 是否输出完整内容（而非截断版本）
    """
    if not table or len(table) == 0:
        return "*空表格*"
    
    md_table = []
    
    # 创建表头
    header = [escape_md_table_cell(str(cell)) for cell in table[0]]
    md_table.append("| " + " | ".join(header) + " |")
    md_table.append("| " + " | ".join(["---"] * len(header)) + " |")
    
    # 添加数据行
    max_rows = len(table) if full_content else min(len(table), 15)  # 限制显示行数
    for row in table[1:max_rows]:
        cells = [escape_md_table_cell(str(cell)) for cell in row]
        md_table.append("| " + " | ".join(cells) + " |")
    
    # 如果表格有更多行且不是完整内容模式
    if len(table) > max_rows and not full_content:
        md_table.append(f"\n*显示前 {max_rows-1} 行数据，共 {len(table)-1} 行*")
    
    return "\n".join(md_table)

# 健康检查接口
@app.route("/health", methods=["GET"])
def health_check():
    """API健康检查"""
    return jsonify({
        "status": "healthy",
        "version": "1.0.0",
        "supported_formats": list(ALLOWED_EXTENSIONS),
        "upload_dir": app.config["UPLOAD_FOLDER"],
        "max_content_length_mb": app.config["MAX_CONTENT_LENGTH"] / (1024 * 1024)
    })

# 专门用于获取Markdown格式的解析结果
@app.route("/files/<path:filename>/markdown", methods=["GET"])
def get_file_markdown(filename):
    """获取特定文件的Markdown格式解析内容"""
    try:
        file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        logger.info(f"请求获取文件Markdown内容: {filename}, 完整路径: {file_path}")
        
        if not os.path.exists(file_path):
            return jsonify({"error": "文件不存在"}), 404
        
        if not os.path.isfile(file_path):
            return jsonify({"error": "路径不是文件"}), 400
		
        # 解析文件
        parsed_data = parse_file(file_path)
        
        # 文件信息
        file_info = {
            "filename": filename,
            "file_type": filename.rsplit(".", 1)[-1].lower() if "." in filename else "",
            "file_size_kb": os.path.getsize(file_path) / 1024,
            "last_modified": os.path.getmtime(file_path)
        }
        
        # 生成Markdown
        markdown_content = convert_to_markdown(parsed_data, file_info)
        
        # 返回Markdown内容
        return Response(markdown_content, mimetype='text/markdown')
    except Exception as e:
        logger.error(f"获取Markdown内容失败: {str(e)}")
        error_md = f"# 错误\n\n生成Markdown内容时出错: {str(e)}"
        return Response(error_md, mimetype='text/markdown')

if __name__ == "__main__":
    logger.info("文档解析API服务启动中...")
    logger.info(f"上传目录: {app.config['UPLOAD_FOLDER']}")
    logger.info(f"支持的文件格式: {', '.join(sorted(ALLOWED_EXTENSIONS))}")
    
    # 检查依赖项安装情况
    logger.info("检查依赖项:")
    logger.info(f"- PIL/Pillow: {'已安装' if has_image_processing else '未安装'}")
    logger.info(f"- python-docx: {'已安装' if has_docx else '未安装'}")
    logger.info(f"- pdfplumber: {'已安装' if has_pdfplumber else '未安装'}")
    logger.info(f"- pytesseract: {'已安装' if has_pytesseract else '未安装'}")
    logger.info(f"- PaddleOCR: {'已安装' if has_paddleocr else '未安装'}")
    logger.info(f"- langchain: {'已安装' if has_langchain else '未安装'}")
    
    logger.info("支持的输出格式: JSON, Markdown")
    app.run(host="0.0.0.0", port=5017, debug=False, use_reloader=False)