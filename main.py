import argparse
import base64
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
import webbrowser
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import img2pdf
import markdown
from mistralai import DocumentURLChunk, Mistral
from mistralai.models import OCRResponse
from PIL import Image

# 添加PyQt5相关导入
try:
    from PyQt5 import QtWidgets
    from PyQt5.QtCore import Qt, QUrl
    from PyQt5.QtWebEngineWidgets import QWebEngineView
    from PyQt5.QtWidgets import QApplication, QWidget

    HAVE_PYQT = True
except ImportError:
    HAVE_PYQT = False

# 支持的图片扩展名列表
SUPPORTED_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".tif", ".webp"}


# 告知用户安装必要的库
def show_install_message(missing_library, additional_info=""):
    return f"""需要安装额外的库

为了获得更好的预览体验，请安装 {missing_library} 库:

pip install {missing_library}

{additional_info}

注意: 安装后需要重启应用程序才能生效。
"""


# 显示简化的安装依赖信息
def show_simple_dependencies_message():
    return """Mistral OCR Markdown Viewer 依赖

为正常运行，需要以下库:
- mistralai, img2pdf, Pillow: 核心OCR功能
- markdown, md4mathjax, markdown-extra: 预览功能
- PyQt5, PyQtWebEngine: 可选的增强预览功能

如有功能问题，请确保所有依赖已正确安装。
"""


def replace_images_in_markdown(markdown_str: str, images_dict: dict) -> str:
    for img_name, img_path in images_dict.items():
        markdown_str = markdown_str.replace(f"![{img_name}]({img_name})", f"![{img_name}]({img_path})")
    return markdown_str


def create_output_directory(original_file_path):
    """创建带时间戳的输出目录"""
    outputs_base = "outputs"
    os.makedirs(outputs_base, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    source_file = Path(original_file_path)
    output_dir = os.path.join(outputs_base, f"{timestamp}_{source_file.stem}")
    os.makedirs(output_dir, exist_ok=True)

    return output_dir


def save_ocr_results(ocr_response: OCRResponse, output_dir: str) -> None:
    images_dir = os.path.join(output_dir, "images")
    os.makedirs(images_dir, exist_ok=True)

    all_markdowns = []
    for page in ocr_response.pages:
        page_images = {}
        for img in page.images:
            img_data = base64.b64decode(img.image_base64.split(",")[1])
            img_path = os.path.join(images_dir, f"{img.id}.png")
            with open(img_path, "wb") as f:
                f.write(img_data)
            page_images[img.id] = f"images/{img.id}.png"

        page_markdown = replace_images_in_markdown(page.markdown, page_images)
        all_markdowns.append(page_markdown)

    markdown_content = "\n\n".join(all_markdowns)
    with open(os.path.join(output_dir, "complete.md"), "w", encoding="utf-8") as f:
        f.write(markdown_content)

    return markdown_content, output_dir


def is_image_file(file_path):
    """检查文件是否为支持的图片格式"""
    ext = os.path.splitext(file_path.lower())[1]
    if ext not in SUPPORTED_IMAGE_EXTENSIONS:
        return False

    try:
        with Image.open(file_path) as img:
            return img.format is not None
    except Exception:
        return False


def get_image_format(file_path):
    """获取图片格式"""
    try:
        with Image.open(file_path) as img:
            return img.format or "未知"
    except Exception:
        return "未知"


def convert_image_to_pdf(image_path, output_dir):
    """将图片转换为PDF并保存到输出目录"""
    pdf_path = os.path.join(output_dir, "converted.pdf")

    try:
        with open(pdf_path, "wb") as f:
            f.write(img2pdf.convert(image_path))
        return pdf_path
    except Exception as e:
        raise Exception(f"图片转换为PDF失败: {str(e)}")


def process_file(file_path: str, api_key: str) -> tuple:
    """处理文件（PDF或图片），返回处理结果"""
    output_dir = create_output_directory(file_path)

    origin_file = os.path.join(output_dir, f"origin{Path(file_path).suffix}")
    shutil.copy2(file_path, origin_file)

    if is_image_file(file_path):
        pdf_path = convert_image_to_pdf(file_path, output_dir)
        file_to_process = pdf_path
    else:
        file_to_process = origin_file

    return process_pdf(file_to_process, api_key, output_dir=output_dir)


def process_pdf(pdf_path: str, api_key: str, output_dir: str = None) -> tuple:
    """处理PDF文件，生成OCR结果"""
    if output_dir is None:
        output_dir = create_output_directory(pdf_path)
        origin_file = os.path.join(output_dir, f"origin{Path(pdf_path).suffix}")
        shutil.copy2(pdf_path, origin_file)

    client = Mistral(api_key=api_key)

    pdf_file = Path(pdf_path)
    if not pdf_file.is_file():
        raise FileNotFoundError(f"文件不存在: {pdf_path}")

    uploaded_file = client.files.upload(
        file={
            "file_name": pdf_file.stem,
            "content": pdf_file.read_bytes(),
        },
        purpose="ocr",
    )

    signed_url = client.files.get_signed_url(file_id=uploaded_file.id, expiry=1)
    pdf_response = client.ocr.process(
        document=DocumentURLChunk(document_url=signed_url.url), model="mistral-ocr-latest", include_image_base64=True
    )

    markdown_content, output_dir = save_ocr_results(pdf_response, output_dir)
    return markdown_content, output_dir


# 简单的Markdown格式化器，用于文本预览
def format_markdown_for_text_preview(markdown_text):
    """
    将Markdown文本转换为简单的格式化文本以供预览
    这个函数实现简单的Markdown到文本的转换，以提供基本的可读性
    """
    lines = markdown_text.split("\n")
    formatted_lines = []

    for line in lines:
        # 处理标题
        if line.startswith("# "):
            formatted_lines.append(f"\n{line[2:].upper()}\n{'=' * len(line[2:])}")
        elif line.startswith("## "):
            formatted_lines.append(f"\n{line[3:].upper()}\n{'-' * len(line[3:])}")
        elif line.startswith("### "):
            formatted_lines.append(f"\n{line[4:]}\n{'~' * len(line[4:])}")
        elif line.startswith("#### "):
            formatted_lines.append(f"\n{line[5:]}")
        # 处理列表
        elif line.strip().startswith("- "):
            formatted_lines.append(f"  • {line.strip()[2:]}")
        elif line.strip().startswith("* "):
            formatted_lines.append(f"  • {line.strip()[2:]}")
        # 处理引用
        elif line.strip().startswith("> "):
            formatted_lines.append(f"  | {line.strip()[2:]}")
        # 处理代码块
        elif line.strip().startswith("```"):
            formatted_lines.append("-" * 40)  # 简单分隔符
        # 处理图片链接 ![alt](url)
        elif "![" in line and "](" in line:
            img_alt = line.split("![")[1].split("]")[0]
            formatted_lines.append(f"[图片: {img_alt}]")
        # 其他行保持不变
        else:
            formatted_lines.append(line)

    return "\n".join(formatted_lines)


# HTML模板创建函数，避免重复代码
def create_html_content(markdown_content, output_dir=None):
    """创建用于预览的HTML内容"""
    output_dir_safe = output_dir.replace(os.sep, "/") if output_dir else ""

    if output_dir:
        processed_markdown = markdown_content.replace("](images/", f"](file://{output_dir_safe}/images/")
    else:
        processed_markdown = markdown_content

    # 使用Python markdown库转换为HTML，使用md4mathjax和extra扩展
    try:
        # 同时使用md4mathjax和extra扩展来支持数学公式和增强的Markdown格式
        html_body = markdown.markdown(processed_markdown, extensions=["md4mathjax", "extra"])
    except (ImportError, ValueError):
        # 如果md4mathjax不可用，尝试仅使用extra
        try:
            html_body = markdown.markdown(processed_markdown, extensions=["extra"])
            print("注意: md4mathjax扩展不可用，数学公式可能无法正确渲染")
        except Exception as e:
            # 如果markdown库或扩展完全不可用，使用简单的HTML转义
            html_body = processed_markdown.replace("<", "&lt;").replace(">", "&gt;")
            print(f"警告: markdown转换出错，使用基本HTML转义。错误信息: {str(e)}")

    # HTML模板 - 包含MathJax支持
    html_template = """<!DOCTYPE html>
<html>
<head>
    <title>Markdown Preview</title>
    <meta charset="utf-8">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/github-markdown-css/github-markdown.min.css">
</head>
<body>
    <div id="content" class="markdown-body">
        HTML_CONTENT_PLACEHOLDER
    </div>
</body>
</html>"""

    # 替换HTML内容
    html_content = html_template.replace("HTML_CONTENT_PLACEHOLDER", html_body)
    return html_content


# 创建Qt的WebView嵌入到Tkinter
class TkinterPyQtBridge:
    """使用PyQt5的WebView创建单独的预览窗口"""

    def __init__(self, parent_window=None, width=800, height=600):
        # 初始化Qt应用程序（如果尚未初始化）
        self.app = QtWidgets.QApplication.instance()
        if not self.app:
            self.app = QtWidgets.QApplication(sys.argv)

        # 创建主窗口
        self.window = QtWidgets.QMainWindow()
        self.window.setWindowTitle("Markdown 预览")
        self.window.resize(width, height)

        # 创建WebView
        self.webview = QWebEngineView()
        self.window.setCentralWidget(self.webview)

        # 记住父窗口以便于定位
        self.parent_window = parent_window

        # 初始化时不显示窗口
        self.is_visible = False

    def load_html(self, html_content):
        """加载HTML内容"""
        self.webview.setHtml(html_content)
        # 如果已经是可见状态，则刷新显示
        if self.is_visible:
            self.show()

    def load_url(self, url):
        """加载URL"""
        self.webview.load(QUrl(url))
        # 如果已经是可见状态，则刷新显示
        if self.is_visible:
            self.show()

    def show(self):
        """显示窗口"""
        # 定位窗口 - 尝试放在父窗口旁边
        if self.parent_window:
            try:
                # 获取父窗口位置和大小
                parent_x = self.parent_window.winfo_x()
                parent_y = self.parent_window.winfo_y()
                parent_width = self.parent_window.winfo_width()

                # 将预览窗口定位在父窗口的右侧
                self.window.move(parent_x + parent_width + 10, parent_y)
            except:
                # 如果无法获取位置，则使用默认位置
                pass

        # 显示窗口
        self.window.show()
        self.window.raise_()  # 确保窗口在最前面
        self.is_visible = True

    def hide(self):
        """隐藏窗口"""
        self.window.hide()
        self.is_visible = False

    def toggle_visibility(self):
        """切换窗口可见性"""
        if self.is_visible:
            self.hide()
        else:
            self.show()


class MarkdownPreviewApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Mistral OCR Markdown Viewer")
        self.geometry("1200x800")

        self.api_key = os.environ.get("MISTRAL_API_KEY", "")

        self.create_widgets()
        self.current_output_dir = None
        self.temp_html_path = None

        # 检查所需库是否安装
        self.check_dependencies()

    def check_dependencies(self):
        """检查必要的库是否已安装"""
        missing_libraries = []

        # 检查markdown库
        try:
            import markdown

            # 检查md4mathjax扩展
            try:
                markdown.markdown("test", extensions=["md4mathjax"])
            except (ImportError, ValueError):
                missing_libraries.append("md4mathjax")

            # 检查extra扩展
            try:
                markdown.markdown("test", extensions=["extra"])
            except (ImportError, ValueError):
                missing_libraries.append("markdown-extra")

        except ImportError:
            missing_libraries.append("markdown")

        # 检查PyQt库（用于嵌入式预览）
        if not HAVE_PYQT:
            # 这是可选的，所以不加入missing_libraries列表
            # 但我们会显示一个推荐
            self.after(1000, self.show_pyqt_recommendation)

        # 如果有缺失的库，显示消息
        if missing_libraries:
            self.after(
                500,
                lambda: messagebox.showwarning(
                    "缺少依赖库",
                    "以下库缺失，这可能会影响某些功能:\n\n"
                    + "\n".join(missing_libraries)
                    + "\n\n请使用pip安装这些库:\n\npip install "
                    + " ".join(missing_libraries),
                ),
            )

    def create_widgets(self):
        # API密钥设置
        self.settings_frame = ttk.Frame(self)
        self.settings_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(self.settings_frame, text="API 密钥:").pack(side=tk.LEFT, padx=5)
        self.api_key_var = tk.StringVar(value=self.api_key)
        self.api_key_entry = ttk.Entry(self.settings_frame, textvariable=self.api_key_var, width=40, show="*")
        self.api_key_entry.pack(side=tk.LEFT, padx=5)

        self.show_key_var = tk.BooleanVar(value=False)
        self.show_key_check = ttk.Checkbutton(
            self.settings_frame, text="显示密钥", variable=self.show_key_var, command=self.toggle_api_key_visibility
        )
        self.show_key_check.pack(side=tk.LEFT, padx=5)

        # 文件选择区域
        self.file_frame = ttk.LabelFrame(self, text="PDF 或图片文件")
        self.file_frame.pack(fill=tk.X, padx=10, pady=5)

        self.file_path_var = tk.StringVar()
        self.file_path_entry = ttk.Entry(self.file_frame, textvariable=self.file_path_var, width=80)
        self.file_path_entry.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)

        self.browse_button = ttk.Button(self.file_frame, text="选择文件", command=self.select_file)
        self.browse_button.pack(side=tk.LEFT, padx=5, pady=5)

        # 文件类型标签
        self.file_type_var = tk.StringVar(value="未选择文件")
        self.file_type_label = ttk.Label(self.file_frame, textvariable=self.file_type_var)
        self.file_type_label.pack(side=tk.LEFT, padx=5, pady=5)

        # 创建一个强调样式
        self.style = ttk.Style()
        self.style.configure("Accent.TButton", font=("Arial", 12, "bold"))

        # 按钮控制区域 - 将所有按钮放在一排
        self.controls_frame = ttk.Frame(self)
        self.controls_frame.pack(fill=tk.X, padx=10, pady=10)

        # 左侧 OCR 按钮 - 突出显示
        self.left_buttons_frame = ttk.Frame(self.controls_frame)
        self.left_buttons_frame.pack(side=tk.LEFT, padx=5)

        self.process_button = ttk.Button(
            self.left_buttons_frame,
            text="开始 OCR 处理",
            command=self.process_file,
            style="Accent.TButton",
            width=20,
        )
        self.process_button.pack(padx=5, pady=5)

        # 中间区域（用于间隔）
        ttk.Frame(self.controls_frame).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 右侧按钮区域
        self.right_buttons_frame = ttk.Frame(self.controls_frame)
        self.right_buttons_frame.pack(side=tk.RIGHT, fill=tk.X)

        # 刷新预览按钮
        self.refresh_button = ttk.Button(self.right_buttons_frame, text="刷新预览", command=self.refresh_preview)
        self.refresh_button.pack(side=tk.LEFT, padx=5)

        # 打开输出文件夹按钮
        self.open_folder_button = ttk.Button(
            self.right_buttons_frame, text="打开输出文件夹", command=self.open_output_folder
        )
        self.open_folder_button.pack(side=tk.LEFT, padx=5)
        self.open_folder_button.config(state=tk.DISABLED)

        # 显示/隐藏预览窗口按钮
        if HAVE_PYQT:
            self.toggle_preview_button = ttk.Button(
                self.right_buttons_frame, text="显示预览窗口", command=self.toggle_preview_window
            )
            self.toggle_preview_button.pack(side=tk.LEFT, padx=5)
            self.toggle_preview_button.config(state=tk.DISABLED)

        # 在浏览器中打开按钮
        self.preview_button = ttk.Button(
            self.right_buttons_frame, text="在浏览器中预览", command=self.preview_in_browser
        )
        self.preview_button.pack(side=tk.LEFT, padx=5)

        # 状态条
        self.status_var = tk.StringVar(value="就绪 - 请选择一个 PDF 或图片文件进行 OCR 处理")
        self.status_bar = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # 进度条
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_bar = ttk.Progressbar(self, variable=self.progress_var, mode="indeterminate")
        self.progress_bar.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)

        # 主内容区域
        self.main_frame = ttk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Markdown源码区域
        self.source_frame = ttk.LabelFrame(self.main_frame, text="Markdown 源码")
        self.source_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Markdown编辑区域
        self.source_text = ScrolledText(self.source_frame, wrap=tk.WORD, width=80, height=20, font=("Courier", 12))
        self.source_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 创建PyQt预览窗口（如果可用）
        if HAVE_PYQT:
            self.web_bridge = TkinterPyQtBridge(parent_window=self, width=800, height=600)

            # 设置初始HTML
            initial_html = """
            <!DOCTYPE html>
            <html>
            <head>
                <style>
                    body {
                        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif;
                        font-size: 16px;
                        line-height: 1.6;
                        margin: 20px;
                        color: #333;
                    }
                    h2 {
                        color: #333;
                        border-bottom: 1px solid #eee;
                        padding-bottom: 5px;
                        font-weight: 600;
                    }
                    p {
                        margin: 16px 0;
                        line-height: 1.6;
                    }
                </style>
            </head>
            <body>
                <h2>Markdown 预览</h2>
                <p>处理文件后将在此显示预览，支持格式化显示、图片和数学公式渲染。</p>
                <p>使用"显示预览窗口"按钮可以随时打开或关闭此预览。</p>
            </body>
            </html>
            """
            self.web_bridge.load_html(initial_html)
        else:
            # 创建提示信息
            ttk.Label(
                self.source_frame,
                text="提示: 安装PyQt5和PyQtWebEngine可获得更好的预览体验",
                font=("Arial", 9, "italic"),
                foreground="gray",
            ).pack(side=tk.BOTTOM, pady=5)

    def show_pyqt_recommendation(self):
        """显示推荐安装PyQt5的消息"""
        messagebox.showinfo(
            "推荐安装",
            "为了获得内嵌的Markdown预览体验，建议安装以下库:\n\n"
            "- PyQt5: 提供基础Qt功能\n"
            "- PyQtWebEngine: 提供网页渲染功能\n\n"
            "可通过以下命令安装:\n"
            "pip install PyQt5 PyQtWebEngine\n\n"
            "安装后重启应用即可激活内嵌的预览功能。",
        )

    def toggle_api_key_visibility(self):
        if self.show_key_var.get():
            self.api_key_entry.config(show="")
        else:
            self.api_key_entry.config(show="*")

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="选择一个文件",
            filetypes=[
                ("所有支持的文件", "*.pdf *.jpg *.jpeg *.png *.gif *.bmp *.tiff *.tif *.webp"),
                ("PDF文件", "*.pdf"),
                ("图片文件", "*.jpg *.jpeg *.png *.gif *.bmp *.tiff *.tif *.webp"),
                ("所有文件", "*.*"),
            ],
        )
        if file_path:
            self.file_path_var.set(file_path)

            # 检查文件类型并更新显示
            if is_image_file(file_path):
                img_format = get_image_format(file_path)
                file_type = f"图片文件 ({img_format})"
                self.file_type_var.set(file_type)
                self.status_var.set(
                    f'已选择{file_type}: {os.path.basename(file_path)} - 点击"开始 OCR 处理"按钮进行处理 (将自动转换为 PDF)'
                )
            else:
                self.file_type_var.set("PDF文件")
                self.status_var.set(f'已选择PDF文件: {os.path.basename(file_path)} - 点击"开始 OCR 处理"按钮进行处理')

    def process_file(self):
        file_path = self.file_path_var.get()
        api_key = self.api_key_var.get()

        if not file_path:
            messagebox.showerror("错误", "请先选择一个文件")
            return

        if not api_key:
            messagebox.showerror("错误", "请提供 Mistral API 密钥")
            return

        # 禁用所有控制按钮
        self.process_button.config(state=tk.DISABLED)
        self.browse_button.config(state=tk.DISABLED)
        self.refresh_button.config(state=tk.DISABLED)
        self.preview_button.config(state=tk.DISABLED)
        self.open_folder_button.config(state=tk.DISABLED)
        if HAVE_PYQT and hasattr(self, "toggle_preview_button"):
            self.toggle_preview_button.config(state=tk.DISABLED)

        # 启动进度条
        self.progress_bar.start(10)

        # 更新状态信息
        is_img = is_image_file(file_path)
        if is_img:
            self.status_var.set(f"正在处理图片: {os.path.basename(file_path)}...")
        else:
            self.status_var.set(f"正在处理 PDF: {os.path.basename(file_path)}...")

        # 在后台线程中处理
        threading.Thread(target=self._process_file_thread, args=(file_path, api_key), daemon=True).start()

    def _process_file_thread(self, file_path, api_key):
        """在后台线程中处理文件的方法"""
        try:
            # 处理文件并获取结果
            markdown_content, output_dir = process_file(file_path, api_key)
            self.current_output_dir = output_dir

            # 在UI线程中更新显示
            self.after(0, lambda: self._update_ui_with_results(markdown_content))
        except Exception as e:
            # 如果发生错误，在UI线程中显示错误消息
            error_message = str(e)
            self.after(0, lambda: messagebox.showerror("处理错误", error_message))
        finally:
            # 无论如何，恢复UI状态
            self.after(0, self._reset_ui_state)

    def _update_ui_with_results(self, markdown_content):
        # 清空并更新源码区域
        self.source_text.delete(1.0, tk.END)
        self.source_text.insert(tk.END, markdown_content)

        # 更新状态
        file_name = os.path.basename(self.file_path_var.get())
        self.status_var.set(f"处理完成 - {file_name} 的结果保存在 {self.current_output_dir}")

        # 更新预览
        self.refresh_preview()

        # 启用预览按钮和控制
        self.refresh_button.config(state=tk.NORMAL)
        self.preview_button.config(state=tk.NORMAL)
        self.open_folder_button.config(state=tk.NORMAL)
        if HAVE_PYQT and hasattr(self, "toggle_preview_button"):
            self.toggle_preview_button.config(state=tk.NORMAL)

        # 清理旧的临时文件（如果有）
        if self.temp_html_path is not None:
            if os.path.exists(self.temp_html_path):
                os.unlink(self.temp_html_path)  # 删除旧的临时文件
            self.temp_html_path = None

    def _reset_ui_state(self):
        # 恢复所有按钮状态
        self.process_button.config(state=tk.NORMAL)
        self.browse_button.config(state=tk.NORMAL)
        self.refresh_button.config(state=tk.NORMAL)
        self.preview_button.config(state=tk.NORMAL)

        # 停止进度条
        self.progress_bar.stop()
        self.progress_var.set(0)

    def refresh_preview(self):
        """刷新Markdown预览"""
        if not hasattr(self, "source_text"):
            return

        # 获取当前Markdown内容
        markdown_content = self.source_text.get(1.0, tk.END)

        if HAVE_PYQT and hasattr(self, "web_bridge"):
            # 如果有PyQt WebEngine，则更新它的内容
            html_content = create_html_content(markdown_content, self.current_output_dir)
            self.web_bridge.load_html(html_content)

            # 自动显示预览窗口
            if self.current_output_dir and not self.web_bridge.is_visible:
                self.web_bridge.show()
                if hasattr(self, "toggle_preview_button"):
                    self.toggle_preview_button.config(text="隐藏预览窗口")

            self.status_var.set("预览已更新 - 使用预览窗口查看格式化内容")
        else:
            # 没有PyQt时，提示用户使用浏览器预览
            self.status_var.set("请使用'在浏览器中预览'按钮查看格式化内容")

    def open_output_folder(self):
        """打开输出文件夹"""
        if not self.current_output_dir or not os.path.exists(self.current_output_dir):
            messagebox.showerror("错误", "没有可打开的输出文件夹")
            return

        # 使用操作系统的文件管理器打开文件夹
        try:
            if os.name == "nt":  # Windows
                os.startfile(self.current_output_dir)
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["open", self.current_output_dir])
            else:  # Linux
                subprocess.run(["xdg-open", self.current_output_dir])
        except Exception as e:
            messagebox.showerror("打开文件夹错误", f"无法打开文件夹: {str(e)}")

    def preview_in_browser(self):
        """在浏览器中预览文档"""
        if not self.current_output_dir:
            messagebox.showerror("错误", "没有可预览的内容")
            return

        # 获取Markdown内容
        markdown_content = self.source_text.get(1.0, tk.END)

        # 创建一个临时HTML文件
        with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as f:
            html_path = f.name

        # 保存临时文件路径
        self.temp_html_path = html_path

        # 创建HTML内容
        html_content = create_html_content(markdown_content, self.current_output_dir)

        # 写入HTML文件
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html_content)

        # 在浏览器中打开
        webbrowser.open("file://" + html_path)

    def toggle_preview_window(self):
        """显示/隐藏预览窗口"""
        if hasattr(self, "web_bridge"):
            self.web_bridge.toggle_visibility()

            # 更新按钮文本
            if self.web_bridge.is_visible:
                self.toggle_preview_button.config(text="隐藏预览窗口")
            else:
                self.toggle_preview_button.config(text="显示预览窗口")

    def show_about(self):
        """显示关于信息"""
        messagebox.showinfo(
            "关于 Mistral OCR Markdown Viewer",
            "Mistral OCR Markdown Viewer 是一个使用 Mistral API 进行 OCR 识别的工具，"
            "支持 PDF 和图片文件，并提供了 Markdown 预览功能。\n\n" + show_simple_dependencies_message() + "\n\n"
            "作者: Zinan Tang\n"
            "版本: 1.0.0",
        )


def main_cli():
    parser = argparse.ArgumentParser(description="使用Mistral API处理PDF或图片文件进行OCR识别")
    parser.add_argument("files", nargs="+", help="要处理的文件路径，可指定多个PDF或图片文件")
    parser.add_argument(
        "--api-key",
        default=os.environ.get("MISTRAL_API_KEY"),
        help="Mistral API密钥，默认从环境变量MISTRAL_API_KEY中获取",
    )
    parser.add_argument(
        "--auto",
        action="store_true",
        help="自动处理所有文件，不需要确认",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="启动图形用户界面",
    )
    args = parser.parse_args()

    if args.gui:
        app = MarkdownPreviewApp()
        app.mainloop()
        return

    if not args.api_key:
        raise ValueError("未提供API密钥，请设置MISTRAL_API_KEY环境变量或使用--api-key参数")

    # 处理所有输入的文件
    for i, file_path in enumerate(args.files):
        try:
            print(f"开始处理文件 {i + 1}/{len(args.files)}: {file_path}")
            markdown_content, output_dir = process_file(file_path, args.api_key)
            print(f"OCR处理完成。结果保存在: {output_dir}")

            # 如果不是最后一个文件且不是自动模式，则询问用户是否继续
            if i < len(args.files) - 1 and not args.auto:
                response = input(f"\n已完成 {file_path} 的处理。按回车键处理下一个文件，或输入'q'退出: ")
                if response.lower() == "q":
                    print("用户选择退出程序。")
                    break
        except Exception as e:
            print(f"处理文件 {file_path} 时出错: {str(e)}")
            if not args.auto and i < len(args.files) - 1:
                response = input(f"\n处理 {file_path} 时出错。按回车键处理下一个文件，或输入'q'退出: ")
                if response.lower() == "q":
                    print("用户选择退出程序。")
                    break


def main():
    # 检测是否有命令行参数，如果没有则启动GUI
    if len(sys.argv) <= 1:
        app = MarkdownPreviewApp()
        app.mainloop()
    else:
        main_cli()


if __name__ == "__main__":
    main()
