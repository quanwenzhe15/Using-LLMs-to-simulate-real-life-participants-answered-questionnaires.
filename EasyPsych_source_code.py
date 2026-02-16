# -*- coding: utf-8 -*-
"""
Questionnaire Simulation System (Adapted for American Participants)
- Reads subject background Excel (Gender/Age/Highest Education Level only)
- Calls Alibaba Cloud Qwen-plus API for simulated responses
- Retains target dimensions: Emotional Abuse, Emotional Neglect, Supervisor Support, Personal Mastery, Perceived Constraints, Job insecurity
- Features: Random question order + No same dimension for 4 consecutive times + API retry + Failure handling + Fatal error stop & save
- Automatically parses scores, handles reverse coding, outputs standardized Excel results
"""
import os
import re
import pandas as pd
from pathlib import Path
import concurrent.futures
import sys

def resource_path(relative_path):
    """获取资源文件的绝对路径，兼容PyInstaller打包后的环境"""
    try:
        # PyInstaller创建临时文件夹并设置_MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # 正常开发环境
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# Show welcome and license agreement window
import tkinter as tk
from tkinter import messagebox, ttk

# Create welcome and license agreement window
def show_welcome_and_license():
    root = tk.Tk()
    root.title("问卷模拟系统 - 欢迎")
    root.geometry("700x600")
    root.resizable(True, True)
    
    # Create main frame
    main_frame = tk.Frame(root, padx=20, pady=20)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # Welcome message
    welcome_label = tk.Label(main_frame, text="欢迎使用EasyPsych问卷模拟系统", font=('Arial', 14, 'bold'))
    welcome_label.pack(pady=10)
    
    # Create a notebook for different sections
    notebook = ttk.Notebook(main_frame)
    notebook.pack(fill=tk.BOTH, expand=True, pady=10)
    
    # Functionality tab
    func_tab = tk.Frame(notebook)
    notebook.add(func_tab, text="功能介绍")
    
    func_text = tk.Text(func_tab, wrap=tk.WORD, padx=10, pady=10)
    func_text.pack(fill=tk.BOTH, expand=True)
    func_text.insert(tk.END, "本系统主要功能：\n\n")
    func_text.insert(tk.END, "1. 支持多种问卷文件格式：Excel (.xlsx, .xls)、CSV (.csv) 和 Word (.docx)\n")
    func_text.insert(tk.END, "2. 自动解析问卷结构，提取题目、维度、计分标准等信息\n")
    func_text.insert(tk.END, "3. 支持随机题目顺序，可限制同一维度连续出现数量\n")
    func_text.insert(tk.END, "4. 集成大模型 API，处理复杂问卷结构\n")
    func_text.insert(tk.END, "5. 模拟被试回答，生成标准化结果文件\n")
    func_text.insert(tk.END, "6. 支持反向计分题处理\n")
    func_text.insert(tk.END, "7. 支持自定义计分规则，可根据需要修改计分标准\n")
    func_text.config(state=tk.DISABLED)
    
    # File format tab
    file_tab = tk.Frame(notebook)
    notebook.add(file_tab, text="文件格式要求")
    
    file_text = tk.Text(file_tab, wrap=tk.WORD, padx=10, pady=10)
    file_text.pack(fill=tk.BOTH, expand=True)
    file_text.insert(tk.END, "文件格式要求：\n\n")
    file_text.insert(tk.END, "问卷文件：\n")
    file_text.insert(tk.END, "- Excel 文件 (.xlsx, .xls)：必须包含以下列：题目ID、题目所属维度、题目内容、计分标准\n")
    file_text.insert(tk.END, "- CSV 文件 (.csv)：必须包含以下列：题目ID、题目所属维度、题目内容、计分标准\n")
    file_text.insert(tk.END, "- Word 文件 (.docx)：按维度分节，维度标题以冒号结尾\n")
    file_text.insert(tk.END, "- 题目格式要求：必须包含来回双引号，例如：4. \"People in my family felt close to each other.\" (R)\n")
    file_text.insert(tk.END, "- 支持的题目格式：\n")
    file_text.insert(tk.END, "  * 标准数字编号：1. \"Question text\" (R)\n")
    file_text.insert(tk.END, "  * 带星号：*1. \"Question text\" (R)\n")
    file_text.insert(tk.END, "  * 项目符号：• \"Question text\" (R)\n")
    file_text.insert(tk.END, "  * 带空格：1 \"Question text\" (R)\n")
    file_text.insert(tk.END, "- 计分规则格式：Coding: 后连接的一句话，以句号结尾，例如：Coding: 1 Never true; 2 Rarely true; 3 Sometimes true; 4 Often true; 5 Very often true. \n")
    file_text.insert(tk.END, "- 支持反向计分标记：(R) 或 (反向)\n")
    file_text.insert(tk.END, "- 支持多种计分范围：1-5、1-7、1-6等（自动识别）\n\n")
    file_text.insert(tk.END, "被试背景文件：\n")
    file_text.insert(tk.END, "- 支持格式：Excel (.xlsx, .xls) 和 CSV (.csv)\n")
    file_text.insert(tk.END, "- 强制要求的列：被试ID、年龄、性别\n")
    file_text.insert(tk.END, "- 其他列会被自动解析并加入到提示中\n")
    file_text.insert(tk.END, "- 空值会被填充为'不适用'，并生成缺失值报告\n")
    file_text.insert(tk.END, "- 如果某列缺失值超过20%，会弹窗告知并让用户选择是否继续\n")
    file_text.config(state=tk.DISABLED)
    
    # Usage guide tab
    guide_tab = tk.Frame(notebook)
    notebook.add(guide_tab, text="使用指南")
    
    guide_text = tk.Text(guide_tab, wrap=tk.WORD, padx=10, pady=10)
    guide_text.pack(fill=tk.BOTH, expand=True)
    guide_text.insert(tk.END, "使用指南：\n\n")
    guide_text.insert(tk.END, "1. 运行脚本后，在欢迎窗口中阅读并同意协议\n")
    guide_text.insert(tk.END, "2. 在设置窗口中配置 API 参数（如需要）\n")
    guide_text.insert(tk.END, "3. 选择是否启用随机题目顺序及连续维度限制\n")
    guide_text.insert(tk.END, "4. 选择问卷文件（Excel、CSV 或 Word 格式）\n")
    guide_text.insert(tk.END, "5. 选择被试背景文件（Excel 格式）\n")
    guide_text.insert(tk.END, "6. 选择输出结果路径\n")
    guide_text.insert(tk.END, "7. 点击'开始运行'按钮开始处理\n")
    guide_text.config(state=tk.DISABLED)
    
    # License tab
    license_tab = tk.Frame(notebook)
    notebook.add(license_tab, text="许可证协议")
    
    license_text = tk.Text(license_tab, wrap=tk.WORD, padx=10, pady=10)
    license_text.pack(fill=tk.BOTH, expand=True)
    license_text.insert(tk.END, "GNU GENERAL PUBLIC LICENSE\n")
    license_text.insert(tk.END, "Version 3, 29 June 2007\n\n")
    license_text.insert(tk.END, "本程序是自由软件：您可以根据自由软件基金会发布的 GNU 通用公共许可证条款\n")
    license_text.insert(tk.END, "（本许可证的第 3 版或您选择的任何更高版本）来重新分发和/或修改它。\n\n")
    license_text.insert(tk.END, "本程序的发布是希望它能有用，但没有任何担保；甚至没有对适销性或\n")
    license_text.insert(tk.END, "特定用途适用性的默示担保。有关详细信息，请参阅 GNU 通用公共许可证。\n\n")
    license_text.insert(tk.END, "您应该已经收到了一份 GNU 通用公共许可证的副本。如果没有，请参见\n")
    license_text.insert(tk.END, "<https://www.gnu.org/licenses/>.\n")
    license_text.config(state=tk.DISABLED)
    
    # Agreement frame
    agree_frame = tk.Frame(main_frame, pady=10)
    agree_frame.pack(fill=tk.X)
    
    agree_var = tk.BooleanVar(value=False)
    agree_checkbox = tk.Checkbutton(agree_frame, text="我已阅读并同意上述条款和条件", variable=agree_var, font=('Arial', 10))
    agree_checkbox.pack(pady=5)
    
    # Button frame
    button_frame = tk.Frame(main_frame, pady=10)
    button_frame.pack(fill=tk.X)
    
    def on_continue():
        if agree_var.get():
            root.destroy()
        else:
            messagebox.showerror("错误", "您必须同意条款和条件才能继续使用本程序")
    
    def on_cancel():
        root.destroy()
        # Exit the program
        import sys
        sys.exit(0)
    
    continue_button = tk.Button(button_frame, text="继续", command=on_continue, font=('Arial', 10, 'bold'), width=15)
    continue_button.pack(side=tk.RIGHT, padx=5)
    
    cancel_button = tk.Button(button_frame, text="退出", command=on_cancel, font=('Arial', 10), width=10)
    cancel_button.pack(side=tk.RIGHT, padx=5)
    
    # Run the window
    root.mainloop()

# Show welcome and license window
show_welcome_and_license()

# Check and install required dependencies with user consent
required_packages = ['tenacity', 'python-docx']
optional_packages = ['pypinyin']  # Optional packages for enhanced functionality
missing_packages = []

# First check for missing packages
for package in required_packages:
    try:
        if package == 'tenacity':
            from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
        elif package == 'python-docx':
            # 先尝试导入，即使失败也不立即报错
            from docx import Document
    except ImportError:
        missing_packages.append(package)

# Check for optional packages
optional_missing = []
for package in optional_packages:
    try:
        if package == 'pypinyin':
                from pypinyin import lazy_pinyin  # type: ignore
    except ImportError:
        optional_missing.append(package)

# If there are missing packages, ask user for consent to install
if missing_packages:
    # Initialize tkinter for the message box
    import tkinter as tk
    from tkinter import messagebox
    
    # Create a root window but hide it
    root = tk.Tk()
    root.withdraw()
    
    # Prepare message
    packages_str = ', '.join(missing_packages)
    message = f"系统检测到缺少以下必要的库：\n{packages_str}\n\n是否同意自动安装这些库？"
    
    # Ask user for consent
    user_consent = messagebox.askyesno("缺少必要库", message)
    
    # Destroy the root window
    root.destroy()
    
    if user_consent:
        # Install missing packages
        for package in missing_packages:
            print(f"Installing required package '{package}'...")
            os.system(f"pip install {package}")
        
        # Re-import after installation
        for package in required_packages:
            if package == 'tenacity':
                from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
            elif package == 'python-docx':
                from docx import Document
    else:
        # User不同意安装，停止运行
        print("用户不同意安装必要的库，程序退出")
        exit(1)

# Check for optional packages
if optional_missing:
    # Initialize tkinter for the message box
    import tkinter as tk
    from tkinter import messagebox
    
    # Create a root window but hide it
    root = tk.Tk()
    root.withdraw()
    
    # Prepare message
    packages_str = ', '.join(optional_missing)
    message = f"系统检测到缺少以下可选的库（用于增强功能）：\n{packages_str}\n\n这些库用于拼音转换功能，缺少它们不会影响基本功能。\n\n是否同意自动安装这些库？"
    
    # Ask user for consent
    user_consent = messagebox.askyesno("缺少可选库", message)
    
    # Destroy the root window
    root.destroy()
    
    if user_consent:
        # Install missing packages
        for package in optional_missing:
            print(f"Installing optional package '{package}'...")
            os.system(f"pip install {package}")

# Now import remaining modules
from openai import OpenAI
from datetime import datetime
# tenacity, docx, tkinter are already imported in the welcome window section
from tkinter import filedialog, messagebox

# ---------------- Core Configuration (Adjust as Needed) ----------------
# API Configuration (Alibaba Cloud Qwen)
def load_config():
    """动态加载配置文件"""
    config_path = resource_path("config.py")
    
    # 创建一个临时模块来加载配置
    import importlib.util
    spec = importlib.util.spec_from_file_location("config", config_path)
    config = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(config)
    
    return config

# Import configuration from config.py
config = load_config()
DASHSCOPE_API_KEY = getattr(config, 'DASHSCOPE_API_KEY', '')
BASE_URL = getattr(config, 'BASE_URL', '')
MODEL_NAME = getattr(config, 'MODEL_NAME', '')

# File Path Configuration
SUBJECT_BACKGROUND_FILE = r"C:\Users\15896\Desktop\我的代码文件\模拟人变量以及相应水平.xlsx"  # Subject background Excel path
OUTPUT_DIR = r"C:\Users\15896\Desktop\我的代码文件"  # Result output directory

MAX_TOKENS = 512  # Maximum length per response
TEMPERATURE = 0.7  # Response diversity (0.7 = close to real human)
#MAX_CONSECUTIVE_SAME_DIM = 3  # Max 3 consecutive questions from same dimension (no 4+)
API_RETRY_TIMES = 3  # API retry times (3 times by default)
API_RETRY_DELAY = 2  # Initial retry delay (2 seconds, exponential backoff)

# DEBUG: 本地测试开关（True=使用模拟 LLM 响应并自动生成测试受试者文件）
DEBUG_MODE = False

# Global flag: Fatal API error (arrearage/access denied)
FATAL_API_ERROR = False
FATAL_ERROR_MSG = ""

# Initialize API Client (OpenAI-compatible format)
client = OpenAI(
    api_key=DASHSCOPE_API_KEY,
    base_url=BASE_URL,
)

# ---------------- File Analysis Logic Explanation ----------------
"""
文件分析逻辑说明：

1. **文件类型检测**：
   - 根据文件扩展名（.xlsx, .xls, .csv, .docx）确定文件类型
   - 调用相应的解析函数

2. **Excel/CSV文件解析**：
   - 读取文件内容
   - 验证必要列：题目ID、题目所属维度、题目内容、计分标准
   - 检查数据完整性，指出具体的无效行
   - 识别反向计分标记（(R)或(反向)）
   - 自动确定计分范围（基于计分标准中的数字）
   - 生成标准化的问题格式

3. **Word文件解析**：
   - 提取文档中的所有文本
   - 识别维度标题（如"Emotional Abuse:", "Emotional Neglect:"等）
   - 提取每个维度下的题目内容
   - 识别计分标准和反向计分标记
   - 生成标准化的问题格式

4. **大模型API集成**：
   - 当正则表达式解析失败时，自动调用大模型API
   - 生成标准化的JSON格式问题列表
   - 处理复杂的问卷结构

5. **动态题目使用**：
   - 完全使用解析得到的题目，而不是硬编码的题目
   - 支持不同维度和题目数量的问卷
   - 动态计算不同维度的量表分数
"""

# Note: 硬编码的QUESTIONS变量已移除，现在完全使用解析得到的题目
# The following QUESTIONS variable is only a placeholder and will be replaced by parsed questions
QUESTIONS = []

# ---------------- Tool Functions ----------------
def load_subject_background(file_path, output_dir, min_age=18, max_age=75):
    """Read subject background Excel/CSV, return standardized subject list"""
    print(f"Reading subject background file: {file_path}")
    print(f"Age range filter: {min_age} - {max_age} years")
    try:
        # Determine file type and read accordingly
        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext == '.csv':
            df = pd.read_csv(file_path, encoding='utf-8-sig')
        else:
            df = pd.read_excel(file_path)
        
        # 调试信息：显示读取的列名
        print(f"读取到的列名: {list(df.columns)}")
        print(f"数据形状: {df.shape}")
        
        # 强制要求的列
        mandatory_cols = ['被试ID', '年龄', '性别']
        missing_mandatory = [col for col in mandatory_cols if col not in df.columns]
        if missing_mandatory:
            raise ValueError(f"背景文件缺少必要列: {', '.join(missing_mandatory)}")
        
        # 调试信息：显示前几行数据
        print("前3行数据:")
        print(df.head(3))
        
        # 1. 统一处理缺失值：把文本「缺失值」替换成NaN，方便后续处理
        df = df.replace("缺失值", pd.NA)
        
        # 2. 记录缺失值位置
        missing_info = []
        for col in df.columns:
            missing_rows = df[df[col].isna()].index.tolist()
            if missing_rows:
                missing_info.append(f"列 '{col}' 在第 {', '.join(map(str, [r+2 for r in missing_rows]))} 行有缺失值")
        
        # 3. 检查列缺失情况
        columns_with_high_missing = []
        total_rows = len(df)
        for col in df.columns:
            missing_count = df[col].isna().sum()
            missing_percentage = (missing_count / total_rows) * 100
            if missing_percentage > 20:
                columns_with_high_missing.append(f"{col} ({missing_percentage:.1f}%)")
        
        # 4. 如果有列缺失超过20%，弹窗告知用户
        if columns_with_high_missing:
            # 使用全局导入的tkinter模块
            import tkinter as tk_local
            from tkinter import messagebox
            
            root = tk_local.Tk()
            root.withdraw()
            
            message = f"以下列的缺失值超过20%：\n{', '.join(columns_with_high_missing)}\n\n是否继续运行？"
            user_choice = messagebox.askyesno("警告：高缺失值", message)
            
            root.destroy()
            
            if not user_choice:
                return []
        
        # 5. 年龄列清洗：转数值类型，检查年龄范围
        df['年龄'] = pd.to_numeric(df['年龄'], errors='coerce').astype('Int64')
        
        # 调试信息：显示年龄列的基本统计
        print(f"年龄列统计:")
        print(f"  非空值数量: {df['年龄'].count()}")
        print(f"  空值数量: {df['年龄'].isna().sum()}")
        print(f"  年龄范围: {df['年龄'].min()} - {df['年龄'].max()}")
        
        # 检查是否有年龄超出范围的被试
        invalid_age_rows = df[(df['年龄'] < min_age) | (df['年龄'] > max_age)]
        
        print(f"年龄无效的被试数量: {len(invalid_age_rows)}")
        
        if not invalid_age_rows.empty:
            # 使用全局导入的tkinter模块
            import tkinter as tk_local
            from tkinter import messagebox
            
            root = tk_local.Tk()
            root.withdraw()
            
            invalid_count = len(invalid_age_rows)
            invalid_ages = invalid_age_rows['年龄'].dropna().unique()
            invalid_ages_str = ', '.join(map(str, sorted(invalid_ages)))
            
            message = f"发现 {invalid_count} 个被试的年龄不在设定范围({min_age}-{max_age}岁)内。\n"
            message += f"无效年龄: {invalid_ages_str}\n\n"
            message += "是否继续运行（将自动过滤掉年龄无效的被试）？"
            
            user_choice = messagebox.askyesno("警告：年龄范围检查", message)
            root.destroy()
            
            if not user_choice:
                return []
        
        # 过滤年龄在范围内的被试
        df = df[(df['年龄'] >= min_age) & (df['年龄'] <= max_age)]
        
        print(f"年龄过滤后的数据形状: {df.shape}")
        
        # 6. 文本列安全处理：先转字符串，再strip
        for col in df.columns:
            if df[col].dtype == 'object' or pd.api.types.is_string_dtype(df[col]):
                # 先转换为字符串，处理可能的编码问题
                df[col] = df[col].astype(str)
                # 填充空值
                df[col] = df[col].fillna("不适用")
                # 去除两端空格
                df[col] = df[col].str.strip()
                
        # 调试信息：显示处理后的数据类型
        print("处理后的数据类型:")
        print(df.dtypes)
        
        # 7. 生成缺失值报告
        if missing_info:
            # 使用全局导入的os模块
            import os as os_module
            report_path = os_module.path.join(output_dir, "missing_values_report.txt")
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write("背景文件缺失值报告\n")
                f.write("=" * 50 + "\n")
                f.write(f"文件路径: {file_path}\n")
                f.write(f"总行数: {total_rows}\n")
                f.write(f"有效行数: {len(df)}\n")
                f.write(f"年龄范围: {min_age}-{max_age}岁\n")
                f.write("\n缺失值位置:\n")
                for info in missing_info:
                    f.write(f"- {info}\n")
                
                # 添加年龄检查信息
                if not invalid_age_rows.empty:
                    f.write(f"\n年龄范围检查:\n")
                    f.write(f"- 过滤掉 {invalid_count} 个年龄无效的被试\n")
                    f.write(f"- 无效年龄: {invalid_ages_str}\n")
            
            print(f"缺失值报告已生成: {report_path}")
        
        # 8. 过滤核心字段全空的行
        df = df.dropna(subset=['性别', '年龄'])
        
        # Convert to subject list
        subjects = []
        for idx, row in df.iterrows():
            subject = {
                "subject_id": int(row['被试ID']) if pd.notna(row['被试ID']) else idx + 1,
                "性别": row['性别'],
                "年龄": row['年龄']
            }
            
            # 添加其他列的信息
            for col in df.columns:
                if col not in mandatory_cols:
                    subject[col] = row[col]
            
            subjects.append(subject)
        
        print(f"最终加载的有效被试数量: {len(subjects)}")
        print(f"每个被试的字段: {list(subjects[0].keys()) if subjects else '无被试'}")
        
        return subjects
    except Exception as e:
        print(f"Failed to read subject background: {str(e)}")
        import traceback
        traceback.print_exc()  
        return []

def generate_subject_prompt(subject, question, column_strategy="保持原样"):
    """Generate subject-specific prompt (English, adapted for American context)"""
    # 优化主管支持备注：根据职业是否为缺失/不适用判断
    supervisor_note = ""
    if "主管支持" in question['dimension']:
        if '职业' in subject and subject['职业'] in ["不适用", "拒绝回答", "不知道"]:
            supervisor_note = " (Note: If you don't have a supervisor or job, answer based on hypothetical work experience or common sense)"
        elif '职业' in subject and '行业' in subject:
            supervisor_note = f" (Note: Answer combined with your occupation as {subject['职业']} in {subject['行业']} industry)"
    
    # Build background information
    background_lines = [
        f"- Gender: {subject['性别']}",
        f"- Age: {subject['年龄']} years old"
    ]
    
    # 自动为所有other fields创建变量并添加到prompt中
    other_fields = []
    for key, value in subject.items():
        if key not in ['subject_id', '性别', '年龄']:
            # 根据用户选择的策略处理字段名
            english_key = process_column_name(key, column_strategy)
            
            # 跳过空值或无效值
            if value not in [None, "", "不适用", "拒绝回答", "不知道"]:
                other_fields.append((english_key, value))
    
    # 添加其他字段到背景信息中
    for english_key, value in other_fields:
        background_lines.append(f"- {english_key}: {value}")
    
    # 动态生成工作相关指导语
    work_guidance = ""
    if any(field[0].lower() in ['occupation', 'industry', 'work'] for field in other_fields):
        work_guidance = "\n4. For work-related questions, answer based on your occupation and industry if applicable;"
    
    # English prompt template with dynamic fields
    prompt = f"""You are a real American citizen with the following personal background:
{chr(10).join(background_lines)}
Fully embody this role, combine American cultural background, life experiences, and true feelings to answer the following questionnaire in the first person{supervisor_note}. Response requirements:
1. Strictly select a score based on the given coding standard (only enter a number between {question['score_range'][0]}-{question['score_range'][1]});
2. Add 1-2 sentences to explain the reason after the score. The reason should match your background and American social culture, avoiding emptiness;
3. Answer naturally and colloquially, like an ordinary American chatting—no formal writing or AI tone;{work_guidance}
5. Do not reveal you are a simulated role, and never say phrases like "as an AI" or "according to the setting";
6. Only answer based on the current task, do not reference any previous responses.
Question: {question['stem']}
Coding Standard: {question['coding']}
Please answer directly without additional formatting."""
    
    # 调试信息：显示生成的prompt结构
    print(f"Generated prompt for subject {subject.get('subject_id', 'unknown')} with {len(other_fields)} additional fields (strategy: {column_strategy})")
    
    return prompt

def process_column_name(column_name, strategy="保持原样"):
    """Process column name based on selected strategy"""    # 标准化列名：移除多余的空格
    normalized_column_name = ''.join(column_name.split())
    
    # 常见字段的英文映射
    field_mappings = {
        '最高教育水平': 'Highest Education Level',
        '职业': 'Occupation', 
        '行业': 'Industry',
        '家庭年总收入': 'Annual Household Income',
        '教育水平': 'Education Level',
        '主管支持': 'Supervisor Support',
        '工作年限': 'Years of Work Experience',
        '婚姻状况': 'Marital Status',
        '居住地': 'Residence',
        '民族': 'Ethnicity',
        '宗教信仰': 'Religious Belief',
        '健康状况': 'Health Status'
    }
    
    # 如果字段在映射表中，使用映射的英文名
    if column_name in field_mappings:
        return field_mappings[column_name]
    if normalized_column_name in field_mappings:
        return field_mappings[normalized_column_name]
    
    # 根据策略处理不在映射表中的字段
    if strategy == "保持原样":
        return column_name  # 保持中文字段名原样
    
    elif strategy == "自动翻译":
        # 简单的中文分词和翻译尝试
        simple_translations = {
            '兴趣': 'Interest', '爱好': 'Hobby', '满意': 'Satisfaction',
            '压力': 'Stress', '生活': 'Life', '工作': 'Work',
            '质量': 'Quality', '水平': 'Level', '程度': 'Degree',
            '关系': 'Relationship', '家庭': 'Family', '社会': 'Social',
            '经济': 'Economic', '心理': 'Psychological', '身体': 'Physical'
        }
        
        # 将中文字段名拆分为单词并尝试翻译
        import re
        words = re.findall(r'[\u4e00-\u9fff]+', column_name)
        translated_words = []
        for word in words:
            if word in simple_translations:
                translated_words.append(simple_translations[word])
            else:
                translated_words.append(word)
        return ' '.join(translated_words)
    
    elif strategy == "拼音转换":
        # 使用拼音作为英文标识
        try:
            from pypinyin import lazy_pinyin  # type: ignore
            return ''.join(lazy_pinyin(column_name))
        except ImportError:
            # 如果pypinyin不可用，使用简单拼音转换
            pinyin_map = {
                'a': 'āáǎà', 'e': 'ēéěè', 'i': 'īíǐì', 'o': 'ōóǒò', 'u': 'ūúǔù', 'v': 'ǖǘǚǜ'
            }
            # 简单的拼音转换（仅处理基本汉字）
            result = ''
            for char in column_name:
                if '\u4e00' <= char <= '\u9fff':  # 中文字符
                    # 简单的拼音映射（实际应用中应该使用pypinyin）
                    result += char
                else:
                    result += char
            return result
    
    elif strategy == "自定义映射":
        # 在编辑提示模板时手动指定，这里保持原样
        return column_name
    
    else:
        # 默认策略：保持原样
        return column_name

def map_text_to_score(text, question):
    """Map text description to score (for responses without explicit numbers)"""
    text_lower = text.lower()
    min_s, max_s = question['score_range']
    coding_type = question['coding']
    
    # 1-5 points (Never true → Very often true)
    if "Never true" in coding_type:
        if any(w in text_lower for w in ["never", "never true", "not at all"]):
            return 1
        elif any(w in text_lower for w in ["rarely", "seldom"]):
            return 2
        elif any(w in text_lower for w in ["sometimes", "occasionally"]):
            return 3
        elif any(w in text_lower for w in ["often", "frequently"]):
            return 4
        elif any(w in text_lower for w in ["very often", "always", "constantly"]):
            return 5
    # 1-5 points (All the time → Never)
    elif "All the time" in coding_type:
        if any(w in text_lower for w in ["all the time", "always"]):
            return 1
        elif any(w in text_lower for w in ["most of the time", "usually"]):
            return 2
        elif any(w in text_lower for w in ["sometimes", "occasionally"]):
            return 3
        elif any(w in text_lower for w in ["rarely", "seldom"]):
            return 4
        elif any(w in text_lower for w in ["never", "not at all"]):
            return 5
    # 1-7 points (Strongly agree → Strongly disagree)
    elif "Strongly agree" in coding_type:
        if any(w in text_lower for w in ["strongly agree", "fully agree", "completely agree"]):
            return 1
        elif any(w in text_lower for w in ["somewhat agree", "partially agree"]):
            return 2
        elif any(w in text_lower for w in ["a little agree", "slightly agree"]):
            return 3
        elif any(w in text_lower for w in ["don't know", "unsure", "no idea"]):
            return 4
        elif any(w in text_lower for w in ["a little disagree", "slightly disagree"]):
            return 5
        elif any(w in text_lower for w in ["somewhat disagree", "partially disagree"]):
            return 6
        elif any(w in text_lower for w in ["strongly disagree", "completely disagree"]):
            return 7
    # 1-5 points (Excellent → Poor)
    elif "Excellent" in coding_type:
        if any(w in text_lower for w in ["excellent", "very good", "definitely"]):
            return 1
        elif any(w in text_lower for w in ["very good", "highly likely"]):
            return 2
        elif any(w in text_lower for w in ["good", "likely"]):
            return 3
        elif any(w in text_lower for w in ["fair", "so-so", "uncertain"]):
            return 4
        elif any(w in text_lower for w in ["poor", "unlikely", "definitely not"]):
            return 5
    
    return None

@retry(
    stop=stop_after_attempt(API_RETRY_TIMES),
    wait=wait_exponential(multiplier=1, min=API_RETRY_DELAY),  # 关键：min=初始延迟，替代错误的initial/initial_delay
    retry=retry_if_exception_type(Exception),
    reraise=True
)
def call_llm(prompt, max_tokens=None):
    """Call Qwen API with retry mechanism, return raw response"""
    global FATAL_API_ERROR, FATAL_ERROR_MSG
    try:
        # Use provided max_tokens or default to MAX_TOKENS
        tokens_to_use = max_tokens if max_tokens is not None else MAX_TOKENS
        
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {"role": "user", "content": prompt}  # 补全你代码截断的messages部分
            ],
            max_tokens=tokens_to_use,
            temperature=TEMPERATURE
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        error_msg = str(e)
        if any(keyword in error_msg for keyword in ["InvalidApiKey", "Arrearage", "AccessDenied"]):
            FATAL_API_ERROR = True
            FATAL_ERROR_MSG = error_msg
        raise

# 如果启用 DEBUG_MODE，则覆盖 call_llm 为本地模拟函数（不调用外部 API）
if 'DEBUG_MODE' in globals() and DEBUG_MODE:
    print("⚙️ DEBUG_MODE 启用：API 调用将被模拟（本地测试）")
    _mock_counter = {'c': 0}
    def _mock_call_llm(prompt):
        # 基于计数循环生成 1-5 的分数，以保证多样性和可预测性
        _mock_counter['c'] += 1
        score = (_mock_counter['c'] % 5) + 1
        reason = f"Mock response #{_mock_counter['c']}: simulated reason matching prompt."
        return f"{score} {reason}"
    # 覆盖真实的 call_llm（用于测试）
    call_llm = _mock_call_llm

def process_single_question(args):
    """Process a single question for a subject (for concurrent execution)"""
    subject, question, column_strategy, api_settings = args
    
    try:
        prompt = generate_subject_prompt(subject, question, column_strategy)
        raw_resp = call_llm(prompt, api_settings.get('max_tokens'))
        score, reason = parse_question_response(raw_resp, question)
        
        # 动态构建响应记录
        response_record = {
            "被试ID": subject['subject_id'],
            "性别": subject['性别'],
            "年龄": subject['年龄'],
            "随机题目序号": question.get('random_index', 0),
            "原始题目ID": question['question_id'],
            "维度": question['dimension'],
            "题目内容（英文）": question['stem'],
            "计分标准（英文）": question['coding'],
            "是否反向计分": question['reverse_coded'],
            "原始响应（英文）": raw_resp,
            "提取分数": score,
            "最终得分": score,
            "回答理由（英文）": reason,
            "作答状态": "成功" if score is not None else "失败"
        }
        
        # 添加所有其他背景文件字段
        for key, value in subject.items():
            if key not in ['subject_id', '性别', '年龄']:
                response_record[key] = value
        
        return response_record, None
        
    except Exception as error_msg:
        # 构建失败记录
        failed_response = {
            "被试ID": subject['subject_id'],
            "性别": subject['性别'],
            "年龄": subject['年龄'],
            "随机题目序号": question.get('random_index', 0),
            "原始题目ID": question['question_id'],
            "维度": question['dimension'],
            "题目内容（英文）": question['stem'],
            "计分标准（英文）": question['coding'],
            "是否反向计分": question['reverse_coded'],
            "原始响应（英文）": f"API_CALL_FAILED: {error_msg}",
            "提取分数": None,
            "最终得分": None,
            "回答理由（英文）": "API call failed",
            "作答状态": "失败"
        }
        
        # 添加所有其他背景文件字段
        for key, value in subject.items():
            if key not in ['subject_id', '性别', '年龄']:
                failed_response[key] = value
        
        return failed_response, {
            "被试ID": subject['subject_id'],
            "题目ID": question['question_id'],
            "错误原因": str(error_msg)
        }

def calculate_scale_scores(responses):
    scale_scores = {}
    # 按维度分组统计分数
    dimension_groups = {}
    for resp in responses:
        dim = resp['维度']
        if dim not in dimension_groups:
            dimension_groups[dim] = []
        if resp['最终得分'] is not None:  # 仅统计有效得分
            dimension_groups[dim].append(resp['最终得分'])
    
    # 动态计算每个维度的分数
    for dimension, scores in dimension_groups.items():
        if scores:
            total_score = sum(scores)
            avg_score = round(total_score / len(scores), 2)
            scale_scores[f'{dimension}_总分'] = total_score
            scale_scores[f'{dimension}_平均分'] = avg_score
        else:
            scale_scores[f'{dimension}_总分'] = None
            scale_scores[f'{dimension}_平均分'] = None
    
    return scale_scores

# ---------------- Parse LLM Response ----------------
def parse_question_response(raw_resp, question):
    """
    Parse the LLM response to extract the score and reason.
    Returns (score, reason).
    """
    # Try to extract the first number in the valid range as the score
    min_s, max_s = question['score_range']
    # Find all numbers in the response
    numbers = re.findall(r'\d+', raw_resp)
    score = None
    for num in numbers:
        n = int(num)
        if min_s <= n <= max_s:
            score = n
            break
    # If not found, try to map text to score
    if score is None:
        score = map_text_to_score(raw_resp, question)
    # Apply reverse coding if needed
    if score is not None and question.get('reverse_coded', False):
        score = max_s + min_s - score
    # Extract reason: remove the score part from the response
    reason = raw_resp
    if score is not None:
        # Remove the score (number) from the start if present
        reason = re.sub(r'^\s*' + str(score) + r'[\s\.\,\:\-]*', '', raw_resp, count=1).strip()
    return score, reason

def get_random_questions(original_questions):
    """Generate random question order with constraint: no same dimension for consecutive times"""
    while True:
        # Create a copy to avoid modifying original list
        random_questions = original_questions.copy()
        import random
        random.shuffle(random_questions)
        
        # Check if constraint is satisfied
        valid = True
        for i in range(len(random_questions) - MAX_CONSECUTIVE_SAME_DIM):
            # Get current dimension and next dimensions
            current_dim = random_questions[i]['dimension']
            consecutive_dims = [random_questions[j]['dimension'] for j in range(i, i + MAX_CONSECUTIVE_SAME_DIM + 1)]
            
            # If all are same dimension, invalid
            if all(dim == current_dim for dim in consecutive_dims):
                valid = False
                break
        
        if valid:
            return random_questions

def save_current_results(all_results, failed_records, out_dir, output_format="xlsx", is_final=False):
    """Save current results immediately (even if process is stopped)"""
    if all_results:
        df_out = pd.DataFrame(all_results)
        
        # 按照被试ID和随机题目序号排序，确保结果顺序正确
        if '被试ID' in df_out.columns and '随机题目序号' in df_out.columns:
            df_out = df_out.sort_values(by=['被试ID', '随机题目序号'])
        
        # 获取所有列
        all_columns = list(df_out.columns)
        
        # 分离背景文件列和系统生成列
        background_columns = []
        system_columns = []
        
        # 系统生成的列（随机题目序号之后的所有列）
        system_generated_columns = [
            "随机题目序号", "原始题目ID", "维度", "题目内容（英文）", "计分标准（英文）", "是否反向计分",
            "原始响应（英文）", "提取分数", "最终得分", "回答理由（英文）", "作答状态"
        ]
        
        # 动态生成的量表分数列
        score_columns = [col for col in df_out.columns if "_总分" in col or "_平均分" in col]
        
        # 将所有列分类
        for col in all_columns:
            if col in system_generated_columns or col in score_columns:
                system_columns.append(col)
            else:
                background_columns.append(col)
        
        # 确保随机题目序号是系统列的第一个
        if "随机题目序号" in system_columns:
            system_columns.remove("随机题目序号")
            system_columns.insert(0, "随机题目序号")
        
        # 组合列顺序：背景文件列 + 系统生成列
        column_order = background_columns + system_columns
        
        # 确保所有列都存在
        for col in column_order:
            if col not in df_out.columns:
                df_out[col] = None
        df_out = df_out[column_order]
        
        # Generate filename
        if is_final:
            # 正常完成时使用固定文件名
            if output_format == "csv":
                output_file = out_dir / f"EasyPsych_Results.csv"
                df_out.to_csv(output_file, index=False, encoding='utf-8-sig')
            else:
                output_file = out_dir / f"EasyPsych_Results.xlsx"
                df_out.to_excel(output_file, index=False, engine='openpyxl')
        else:
            # 中断时使用带时间戳的文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            if output_format == "csv":
                output_file = out_dir / f"Interrupted_Results_{timestamp}.csv"
                df_out.to_csv(output_file, index=False, encoding='utf-8-sig')
            else:
                output_file = out_dir / f"Interrupted_Results_{timestamp}.xlsx"
                df_out.to_excel(output_file, index=False, engine='openpyxl')
        
        print(f"\n Current results saved to: {output_file}")
        
        # Save failed records if any
        if failed_records:
            df_failed = pd.DataFrame(failed_records)
            if output_format == "csv":
                failed_file = out_dir / f"Interrupted_Failed_Records_{timestamp}.csv"
                df_failed.to_csv(failed_file, index=False, encoding='utf-8-sig')
            else:
                failed_file = out_dir / f"Interrupted_Failed_Records_{timestamp}.xlsx"
                df_failed.to_excel(failed_file, index=False, engine='openpyxl')
            print(f" Failed records saved to: {failed_file}")
        
        # Save fatal error info if exists
        if FATAL_API_ERROR:
            error_info = pd.DataFrame([{
                "终止原因": "API致命错误",
                "错误详情": FATAL_ERROR_MSG,
                "终止时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "已处理被试数": len(set([r['被试ID'] for r in all_results])),
                "已处理题目数": len(all_results)
            }])
            if output_format == "csv":
                error_file = out_dir / f"Fatal_Error_Info_{timestamp}.csv"
                error_info.to_csv(error_file, index=False, encoding='utf-8-sig')
            else:
                error_file = out_dir / f"Fatal_Error_Info_{timestamp}.xlsx"
                error_info.to_excel(error_file, index=False, engine='openpyxl')
            print(f"✅ Fatal error info saved to: {error_file}")
    else:
        print("\n⚠️ No results to save (all_results is empty)")

# ---------------- Integrated GUI Settings ----------------
def show_settings_gui():
    """Show integrated settings GUI with API settings, file selection, and options"""
    root = tk.Tk()
    root.title("问卷模拟系统设置")
    root.geometry("600x700")
    root.resizable(True, True)
    
    # Create main frame
    main_frame = tk.Frame(root, padx=20, pady=20)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # Create notebook (tabbed interface)
    notebook = ttk.Notebook(main_frame)
    notebook.pack(fill=tk.BOTH, expand=True)
    
    # ---------------- API Settings Tab ----------------
    api_tab = tk.Frame(notebook)
    notebook.add(api_tab, text="API设置")
    
    # API Key
    tk.Label(api_tab, text="API密钥:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky=tk.W, pady=5, padx=10)
    api_key_var = tk.StringVar(value=DASHSCOPE_API_KEY)
    api_key_entry = tk.Entry(api_tab, textvariable=api_key_var, width=50)
    api_key_entry.grid(row=0, column=1, pady=5, padx=10)
    
    # Base URL
    tk.Label(api_tab, text="基础URL:", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky=tk.W, pady=5, padx=10)
    base_url_var = tk.StringVar(value=BASE_URL)
    base_url_entry = tk.Entry(api_tab, textvariable=base_url_var, width=50)
    base_url_entry.grid(row=1, column=1, pady=5, padx=10)
    
    # Model Name
    tk.Label(api_tab, text="模型名称:", font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky=tk.W, pady=5, padx=10)
    model_name_var = tk.StringVar(value=MODEL_NAME)
    model_name_entry = tk.Entry(api_tab, textvariable=model_name_var, width=50)
    model_name_entry.grid(row=2, column=1, pady=5, padx=10)
    
    # ---------------- Questionnaire Settings Tab ----------------
    q_settings_tab = tk.Frame(notebook)
    notebook.add(q_settings_tab, text="问卷设置")
    
    # Random Question Order
    random_order_var = tk.BooleanVar(value=False)
    tk.Checkbutton(q_settings_tab, text="启用随机题目顺序", variable=random_order_var, font=('Arial', 10)).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=5, padx=10)
    
    # Max Consecutive Same Dimension
    tk.Label(q_settings_tab, text="同一维度最大连续出现数量:", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky=tk.W, pady=5, padx=10)
    max_consecutive_var = tk.IntVar(value=3)
    max_consecutive_spin = tk.Spinbox(q_settings_tab, from_=1, to=10, textvariable=max_consecutive_var, width=10)
    max_consecutive_spin.grid(row=1, column=1, sticky=tk.W, pady=5, padx=10)
    
    # Token limit for API analysis
    tk.Label(q_settings_tab, text="API分析最大token数:", font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky=tk.W, pady=5, padx=10)
    token_limit_var = tk.IntVar(value=4000)
    token_frame = tk.Frame(q_settings_tab)
    token_frame.grid(row=2, column=1, pady=5, padx=10, sticky=tk.W)
    token_scale = tk.Scale(token_frame, from_=1000, to=8000, orient=tk.HORIZONTAL, variable=token_limit_var, 
             length=200, resolution=500)
    token_scale.pack(side=tk.LEFT)
    
    # Update token label when scale changes
    def update_token_label(*args):
        token_label.config(text=f"{token_limit_var.get()} tokens")
    
    token_limit_var.trace("w", update_token_label)
    token_label = tk.Label(token_frame, text=f"{token_limit_var.get()} tokens")
    token_label.pack(side=tk.LEFT, padx=10)
    
    # MAX_TOKENS setting for individual responses
    tk.Label(q_settings_tab, text="单次回答最大token数:", font=('Arial', 10, 'bold')).grid(row=3, column=0, sticky=tk.W, pady=5, padx=10)
    max_tokens_var = tk.IntVar(value=512)
    max_tokens_frame = tk.Frame(q_settings_tab)
    max_tokens_frame.grid(row=3, column=1, pady=5, padx=10, sticky=tk.W)
    tk.Scale(max_tokens_frame, from_=100, to=2000, orient=tk.HORIZONTAL, variable=max_tokens_var, 
             length=200, resolution=50).pack(side=tk.LEFT)
    
    def update_max_tokens_label(*args):
        max_tokens_label.config(text=f"{max_tokens_var.get()} tokens")
    
    max_tokens_var.trace("w", update_max_tokens_label)
    max_tokens_label = tk.Label(max_tokens_frame, text=f"{max_tokens_var.get()} tokens")
    max_tokens_label.pack(side=tk.LEFT, padx=10)
    
    # Age range settings
    tk.Label(q_settings_tab, text="被试年龄范围:", font=('Arial', 10, 'bold')).grid(row=4, column=0, sticky=tk.W, pady=5, padx=10)
    age_frame = tk.Frame(q_settings_tab)
    age_frame.grid(row=4, column=1, pady=5, padx=10, sticky=tk.W)
    
    tk.Label(age_frame, text="最小年龄:").pack(side=tk.LEFT)
    min_age_var = tk.IntVar(value=18)
    min_age_spin = tk.Spinbox(age_frame, from_=0, to=100, textvariable=min_age_var, width=5)
    min_age_spin.pack(side=tk.LEFT, padx=5)
    
    tk.Label(age_frame, text="最大年龄:").pack(side=tk.LEFT, padx=(10, 0))
    max_age_var = tk.IntVar(value=75)
    max_age_spin = tk.Spinbox(age_frame, from_=0, to=100, textvariable=max_age_var, width=5)
    max_age_spin.pack(side=tk.LEFT, padx=5)
    
    # Validate age range
    def validate_age_range():
        min_age = min_age_var.get()
        max_age = max_age_var.get()
        if min_age >= max_age:
            messagebox.showerror("错误", "最小年龄必须小于最大年龄")
            min_age_var.set(18)
            max_age_var.set(75)
        elif min_age < 0 or max_age > 100:
            messagebox.showerror("错误", "年龄范围必须在0-100岁之间")
            min_age_var.set(18)
            max_age_var.set(75)
    
    min_age_var.trace("w", lambda *args: validate_age_range())
    max_age_var.trace("w", lambda *args: validate_age_range())
    
    # Custom scoring rules
    tk.Label(q_settings_tab, text="自定义计分规则:", font=('Arial', 10, 'bold')).grid(row=5, column=0, sticky=tk.W, pady=5, padx=10)
    
    def edit_scoring_rules():
        # Create scoring rules edit window
        scoring_window = tk.Toplevel(root)
        scoring_window.title("编辑计分规则")
        scoring_window.geometry("800x600")
        scoring_window.resizable(True, True)
        
        # Create main frame
        scoring_frame = tk.Frame(scoring_window, padx=20, pady=20)
        scoring_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scoring rules text
        scoring_label = tk.Label(scoring_frame, text="计分规则设置:", font=('Arial', 10, 'bold'))
        scoring_label.pack(pady=10)
        
        # Text widget for scoring rules
        scoring_text = tk.Text(scoring_frame, wrap=tk.WORD, font=('Arial', 10))
        scoring_text.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Insert default scoring rules with examples
        default_rules = "# 计分规则设置\n\n"
        default_rules += "## 1. 计分范围识别规则\n"
        default_rules += "- 自动识别计分范围：根据计分标准中的数字确定\n"
        default_rules += "  例如：包含 '7' 则为 1-7 点计分，包含 '6' 则为 1-6 点计分\n"
        default_rules += "  默认为 1-5 点计分\n\n"
        
        default_rules += "## 2. 反向计分规则\n"
        default_rules += "- 自动识别反向计分标记：(R) 或 (反向)\n"
        default_rules += "- 反向计分计算：(最小值 + 最大值) - 原始分数\n"
        default_rules += "  例如：5点计分中，原始分数为 1，则反向计分为 5\n"
        default_rules += "  例如：7点计分中，原始分数为 2，则反向计分为 6\n\n"
        
        default_rules += "## 3. 维度分数计算规则\n"
        default_rules += "- 维度分数 = 该维度下所有题目分数的总和\n"
        default_rules += "- 支持缺失值处理：仅一个缺失值时使用均值替换\n"
        default_rules += "- 无缺失值时直接求和\n\n"
        
        default_rules += "## 4. 自定义规则示例\n"
        default_rules += "# 示例1：修改计分范围识别\n"
        default_rules += "# score_range = (1, 5)  # 强制使用5点计分\n\n"
        
        default_rules += "# 示例2：修改反向计分计算\n"
        default_rules += "# reverse_score = max_score - (original_score - min_score)\n\n"
        
        default_rules += "# 示例3：修改维度分数计算为平均值\n"
        default_rules += "# dimension_score = sum(scores) / len(scores)\n\n"
        
        scoring_text.insert(tk.END, default_rules)
        
        # Button frame
        button_frame = tk.Frame(scoring_frame, pady=10)
        button_frame.pack(fill=tk.X)
        
        def save_scoring_rules():
            # Get the edited scoring rules
            edited_rules = scoring_text.get(1.0, tk.END).strip()
            # Here you could save the rules to a file or update a global variable
            print("Scoring rules updated:")
            print(edited_rules)
            # Close the window
            scoring_window.destroy()
        
        save_button = tk.Button(button_frame, text="保存", command=save_scoring_rules, font=('Arial', 10, 'bold'), width=15)
        save_button.pack(side=tk.RIGHT, padx=5)
        
        cancel_button = tk.Button(button_frame, text="取消", command=scoring_window.destroy, font=('Arial', 10), width=10)
        cancel_button.pack(side=tk.RIGHT, padx=5)
    
    edit_scoring_button = tk.Button(q_settings_tab, text="编辑计分规则", command=edit_scoring_rules, font=('Arial', 10))
    edit_scoring_button.grid(row=5, column=1, sticky=tk.W, pady=5, padx=10)
    
    # New column name handling strategy
    tk.Label(q_settings_tab, text="新列名处理策略:", font=('Arial', 10, 'bold')).grid(row=6, column=0, sticky=tk.W, pady=5, padx=10)
    column_strategy_var = tk.StringVar(value="保持原样")
    strategy_frame = tk.Frame(q_settings_tab)
    strategy_frame.grid(row=6, column=1, pady=5, padx=10, sticky=tk.W)
    
    # Strategy options with descriptions
    strategies = [
        ("保持原样", "保持中文字段名原样（推荐，AI能理解）"),
        ("自动翻译", "尝试自动翻译为英文"),
        ("拼音转换", "使用拼音作为英文标识"),
        ("自定义映射", "在编辑提示模板时手动指定英文名")
    ]
    
    for i, (strategy, description) in enumerate(strategies):
        rb = tk.Radiobutton(strategy_frame, text=strategy, variable=column_strategy_var, value=strategy)
        rb.grid(row=i, column=0, sticky=tk.W)
        desc_label = tk.Label(strategy_frame, text=description, font=('Arial', 8), fg='gray')
        desc_label.grid(row=i, column=1, sticky=tk.W, padx=(5, 0))
    
    # ---------------- File Selection Tab ----------------
    file_tab = tk.Frame(notebook)
    notebook.add(file_tab, text="文件选择")
    
    # Questionnaire File
    tk.Label(file_tab, text="问卷文件:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky=tk.W, pady=5, padx=10)
    questionnaire_file_var = tk.StringVar()
    questionnaire_entry = tk.Entry(file_tab, textvariable=questionnaire_file_var, width=40)
    questionnaire_entry.grid(row=0, column=1, pady=5, padx=10)
    tk.Button(file_tab, text="浏览", command=lambda: questionnaire_file_var.set(filedialog.askopenfilename(
        title="选择问卷文件",
        filetypes=[("Excel文件", "*.xlsx;*.xls"), ("CSV文件", "*.csv"), ("Word文件", "*.docx"), ("所有文件", "*.*")]
    ))).grid(row=0, column=2, pady=5, padx=10)
    
    # Subject Background File
    tk.Label(file_tab, text="被试背景文件:", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky=tk.W, pady=5, padx=10)
    subject_file_var = tk.StringVar()
    subject_entry = tk.Entry(file_tab, textvariable=subject_file_var, width=40)
    subject_entry.grid(row=1, column=1, pady=5, padx=10)
    tk.Button(file_tab, text="浏览", command=lambda: subject_file_var.set(filedialog.askopenfilename(
        title="选择被试背景文件",
        filetypes=[("Excel文件", "*.xlsx;*.xls"), ("CSV文件", "*.csv"), ("所有文件", "*.*")]
    ))).grid(row=1, column=2, pady=5, padx=10)
    
    # Output Directory
    tk.Label(file_tab, text="输出结果路径:", font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky=tk.W, pady=5, padx=10)
    output_dir_var = tk.StringVar(value=OUTPUT_DIR)
    output_entry = tk.Entry(file_tab, textvariable=output_dir_var, width=40)
    output_entry.grid(row=2, column=1, pady=5, padx=10)
    tk.Button(file_tab, text="浏览", command=lambda: output_dir_var.set(filedialog.askdirectory(
        title="选择输出结果路径"
    ))).grid(row=2, column=2, pady=5, padx=10)
    
    # Output Format
    tk.Label(file_tab, text="输出格式:", font=('Arial', 10, 'bold')).grid(row=3, column=0, sticky=tk.W, pady=5, padx=10)
    output_format_var = tk.StringVar(value="xlsx")
    format_frame = tk.Frame(file_tab)
    format_frame.grid(row=3, column=1, pady=5, padx=10, sticky=tk.W)
    tk.Radiobutton(format_frame, text="Excel (.xlsx)", variable=output_format_var, value="xlsx").pack(side=tk.LEFT, padx=10)
    tk.Radiobutton(format_frame, text="CSV (.csv)", variable=output_format_var, value="csv").pack(side=tk.LEFT, padx=10)
    
    # ---------------- Prompt Edit Button ----------------
    def edit_prompt():
        # Create prompt edit window
        prompt_window = tk.Toplevel(root)
        prompt_window.title("编辑提示模板")
        
        # Get screen resolution and adjust window size
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        window_width = int(screen_width * 0.7)
        window_height = int(screen_height * 0.7)
        
        prompt_window.geometry(f"{window_width}x{window_height}")
        prompt_window.resizable(True, True)
        
        # Create main frame
        prompt_frame = tk.Frame(prompt_window, padx=20, pady=20)
        prompt_frame.pack(fill=tk.BOTH, expand=True)
        
        # Prompt text
        prompt_label = tk.Label(prompt_frame, text="提示模板:", font=('Arial', 10, 'bold'))
        prompt_label.pack(pady=5)
        
        # Text widget for prompt editing
        prompt_text = tk.Text(prompt_frame, wrap=tk.WORD, font=('Arial', 10))
        prompt_text.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Insert default prompt
        default_prompt = """You are a real American citizen with the following personal background:
{background_info}
Fully embody this role, combine American cultural background, life experiences, and true feelings to answer the following questionnaire in the first person{supervisor_note}. Response requirements:
1. Strictly select a score based on the given coding standard (only enter a number between {score_range});
2. Add 1-2 sentences to explain the reason after the score. The reason should match your background and American social culture, avoiding emptiness;
3. Answer naturally and colloquially, like an ordinary American chatting—no formal writing or AI tone;
4. For work-related questions, answer based on your occupation and industry if applicable;
5. Do not reveal you are a simulated role, and never say phrases like "as an AI" or "according to the setting";
6. Only answer based on the current task, do not reference any previous responses.
Question: {question_stem}
Coding Standard: {coding_standard}
Please answer directly without additional formatting."""
        prompt_text.insert(tk.END, default_prompt)
        
        # Button frame
        button_frame = tk.Frame(prompt_frame, pady=10)
        button_frame.pack(fill=tk.X)
        
        def save_prompt():
            # Get the edited prompt
            edited_prompt = prompt_text.get(1.0, tk.END).strip()
            # Here you could save the prompt to a file or update a global variable
            print("Prompt updated:")
            print(edited_prompt)
            # Close the window
            prompt_window.destroy()
        
        save_button = tk.Button(button_frame, text="保存", command=save_prompt, font=('Arial', 10, 'bold'), width=15)
        save_button.pack(side=tk.RIGHT, padx=5)
        
        cancel_button = tk.Button(button_frame, text="取消", command=prompt_window.destroy, font=('Arial', 10), width=10)
        cancel_button.pack(side=tk.RIGHT, padx=5)
    
    # ---------------- Submit Button ----------------
    def submit():
        # Validate inputs
        if not questionnaire_file_var.get():
            messagebox.showerror("错误", "未选择问卷文件")
            return
        if not subject_file_var.get():
            messagebox.showerror("错误", "未选择被试背景文件")
            return
        if not output_dir_var.get():
            messagebox.showerror("错误", "未选择输出结果路径")
            return
        
        # Close window and return values
        root.destroy()
        
    # Button frame
    button_frame = tk.Frame(main_frame, pady=10)
    button_frame.pack(fill=tk.X)
    
    # Prompt edit button
    tk.Button(button_frame, text="编辑提示模板", command=edit_prompt, font=('Arial', 10), width=15).pack(side=tk.LEFT, padx=5)
    
    # Submit button
    tk.Button(button_frame, text="开始运行", command=submit, font=('Arial', 10, 'bold'), width=20).pack(side=tk.RIGHT, padx=5)
    
    # Run the GUI
    root.mainloop()
    
    # Return all settings
    return {
        'api_key': api_key_var.get(),
        'base_url': base_url_var.get(),
        'model_name': model_name_var.get(),
        'random_order': random_order_var.get(),
        'max_consecutive': max_consecutive_var.get(),
        'token_limit': token_limit_var.get(),
        'max_tokens': max_tokens_var.get(),
        'min_age': min_age_var.get(),
        'max_age': max_age_var.get(),
        'column_strategy': column_strategy_var.get(),
        'questionnaire_file': questionnaire_file_var.get(),
        'subject_file': subject_file_var.get(),
        'output_dir': output_dir_var.get(),
        'output_format': output_format_var.get()
    }

# ---------------- Questionnaire File Parser ----------------
def parse_questionnaire_file(file_path, token_limit=4000):
    """Parse questionnaire file based on file type"""
    file_ext = os.path.splitext(file_path)[1].lower()
    
    if file_ext in ['.xlsx', '.xls', '.csv']:
        return parse_excel_csv_questionnaire(file_path)
    elif file_ext == '.docx':
        return parse_word_questionnaire(file_path, token_limit)
    else:
        messagebox.showerror("错误", f"不支持的文件格式: {file_ext}")
        return None

def parse_excel_csv_questionnaire(file_path):
    """Parse Excel/CSV questionnaire file"""
    try:
        # Determine file type and read accordingly
        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext == '.csv':
            df = pd.read_csv(file_path, encoding='utf-8-sig')
        else:
            df = pd.read_excel(file_path)
        
        # Required columns
        required_cols = ['题目ID', '题目所属维度', '题目内容', '计分标准']
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            messagebox.showerror("错误", f"问卷文件缺少必要列: {', '.join(missing_cols)}")
            return None
        
        # Validate data integrity
        invalid_rows = []
        for idx, row in df.iterrows():
            row_num = idx + 2  # Excel rows start at 1, plus header
            missing_values = []
            for col in required_cols:
                if pd.isna(row[col]) or str(row[col]).strip() == '':
                    missing_values.append(col)
            if missing_values:
                invalid_rows.append(f"第{row_num}行: 缺少{', '.join(missing_values)}")
        
        if invalid_rows:
            error_msg = "发现以下无效行:\n" + "\n".join(invalid_rows)
            messagebox.showerror("错误", error_msg)
            return None
        
        # Parse questions
        questions = []
        for idx, row in df.iterrows():
            question_id = str(row['题目ID']).strip()
            dimension = str(row['题目所属维度']).strip()
            stem = str(row['题目内容']).strip()
            coding = str(row['计分标准']).strip()
            
            # Determine reverse coding (check if '(R)' is in stem)
            reverse_coded = '(R)' in stem or '(反向)' in stem
            if reverse_coded:
                # Remove (R) marker from stem
                stem = stem.replace('(R)', '').replace('(反向)', '').strip()
            
            # Determine score range from coding
            score_range = (1, 5)  # Default
            if '7' in coding:
                score_range = (1, 7)
            
            questions.append({
                "question_id": question_id,
                "dimension": dimension,
                "stem": stem,
                "coding": coding,
                "reverse_coded": reverse_coded,
                "score_range": score_range
            })
        
        print(f"Successfully parsed {len(questions)} questions from Excel/CSV file")
        return questions
        
    except Exception as e:
        messagebox.showerror("错误", f"解析问卷文件失败: {str(e)}")
        print(f"Error parsing Excel/CSV file: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

def parse_word_questionnaire(file_path, token_limit=4000):
    """Parse Word questionnaire file"""
    try:
        doc = Document(file_path)
        full_text = []
        
        # Extract all text from the document
        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                full_text.append(text)
        
        # Join all text into a single string
        document_text = '\n'.join(full_text)
        
        print(f"\n=== Word文档解析调试信息 ===")
        print(f"读取到 {len(full_text)} 行文本")
        
        # Split document into sections based on dimension headers
        # Look for patterns like "Emotional Abuse:", "Emotional Neglect:", etc.
        sections = []
        current_section = []
        dimension_names = []
        
        # 常见维度名称列表（用于加强识别）
        known_dimensions = [
            'Emotional Abuse',
            'Emotional Neglect', 
            'Supervisory Support',
            'Perceived Control',
            'Personal Mastery',
            'Perceived Constraints',
            'Job Insecurity'
        ]
        
        print(f"\n开始识别维度标题...")
        for line_idx, line in enumerate(full_text):
            # Check if this line is a dimension header
            is_dimension = False
            
            # Check if line ends with colon (allowing spaces before colon) and isn't a known non-dimension line
            # 对空格更加宽容，允许冒号前后有空格，同时支持全角冒号
            stripped_line = line.strip()
            # 检查是否以冒号结尾（支持半角和全角冒号）
            ends_with_colon = stripped_line.endswith(':') or stripped_line.endswith('：')
            # 检查是否不是非维度行
            is_not_non_dimension = not any(prefix in line for prefix in ['Items:', 'Question', 'Coding:', 'Scaling:', 'Scoring Key:'])
            # 检查是否不是空行
            is_not_empty = len(stripped_line) > 0
            # 检查是否不是只有冒号的行
            is_not_only_colon = not (stripped_line == ':' or stripped_line == '：')
            
            if ends_with_colon and is_not_non_dimension and is_not_empty and is_not_only_colon:
                is_dimension = True
                print(f"  行 {line_idx+1}: 识别为维度（冒号结尾）: '{line}'")
            
            # 检查是否是已知的维度名称（即使不带冒号）
            else:
                # 只有当行长度较短且看起来像维度标题时才检查
                # 避免把包含"Personal Mastery"的长描述行也识别为维度
                if len(stripped_line) < 50:  # 限制行长度，避免识别描述性文字
                    for known_dim in known_dimensions:
                        if known_dim.lower() in stripped_line.lower():
                            # 确保不是coding、scaling等行
                            if not any(prefix in line for prefix in ['Items:', 'Question', 'Coding:', 'Scaling:', 'Scoring Key:']):
                                is_dimension = True
                                print(f"  行 {line_idx+1}: 识别为维度（已知名称）: '{line}' (匹配: {known_dim})")
                                break
            
            if is_dimension:
                if current_section:
                    sections.append(current_section)
                current_section = [line]
                # Extract dimension name
                # 处理半角和全角冒号
                dim_name = line.rstrip(':').rstrip('：').strip()
                # Remove trailing "scale" or "Scale" if present
                dim_name = re.sub(r'\s*(scale|Scale)$', '', dim_name, flags=re.IGNORECASE)
                # Ensure consistent spacing
                dim_name = ' '.join(dim_name.split())
                # 特殊处理常见维度名称，确保正确识别
                common_dimensions = {
                    'Emotional Abuse': 'Emotional Abuse',
                    'Emotional Neglect': 'Emotional Neglect',
                    'Supervisory Support': 'Supervisory Support',
                    'Perceived Control': 'Perceived Control',
                    'Personal Mastery': 'Personal Mastery',
                    'Perceived Constraints': 'Perceived Constraints',
                    'Job Insecurit': 'Job Insecurity',
                    'Job Insecurity': 'Job Insecurity'
                }
                # 标准化维度名称
                if dim_name in common_dimensions:
                    dim_name = common_dimensions[dim_name]
                # 修复常见拼写错误
                if dim_name == 'Job Insecurit':
                    dim_name = 'Job Insecurity'
                # 如果维度名称不在映射表中，直接使用原始文字部分
                # 确保维度名称不为空
                if not dim_name:
                    dim_name = f"维度{len(dimension_names) + 1}"
                dimension_names.append(dim_name)
            else:
                current_section.append(line)
        
        # Add the last section
        if current_section:
            sections.append(current_section)
        
        print(f"\n识别到 {len(sections)} 个section")
        print(f"识别到的维度名称: {dimension_names}")
        
        if not sections:
            messagebox.showerror("错误", "Word文件中未找到问卷维度")
            return None
        
        # Parse each section
        questions = []
        dimension_counter = {}
        
        for section_idx, section in enumerate(sections):
            section_text = '\n'.join(section)
            dimension = dimension_names[section_idx] if section_idx < len(dimension_names) else f"维度{section_idx+1}"
            print(f"\n处理Section {section_idx+1}, 维度: {dimension}")
            print(f"Section内容行数: {len(section)}")
            
            # Extract coding information with enhanced patterns
            # 按照用户要求：计分规则识别为'Coding:(后连接的一句话).'
            coding_patterns = [
                r'Coding:\s*(.*?)\.',  # 用户要求的格式：Coding: 后连接的一句话，以句号结尾
                r'(?:Scoring Key|Scoring):\s*(.*?)\.',  # 其他计分规则格式
                r'(?:Responses are obtained using a|Scoring Key:)\s*(.*?)(?:\n|$)',
                r'(?:1 = |Strongly agree;).*?(?:\n|$)',  # Additional patterns for coding
                r'(?:Scoring Key:|Coding:).*?(?:1.*?5|1.*?7)(?:\n|$)'
            ]
            
            coding = "1-5 Likert scale"  # Default
            for pattern in coding_patterns:
                coding_match = re.search(pattern, section_text, re.DOTALL | re.IGNORECASE)
                if coding_match:
                    coding = coding_match.group(1).strip() if len(coding_match.groups()) > 0 else coding_match.group(0).strip()
                    break
            
            # Determine score range with enhanced logic
            score_range = (1, 5)  # Default
            if '7' in coding:
                score_range = (1, 7)
            elif '6' in coding:
                score_range = (1, 6)
            elif '4' in coding:
                score_range = (1, 4)
            
            # Extract questions with enhanced regex patterns
            # 根据用户要求：题目必须包含来回双引号，例如：4. "People in my family felt close to each other." (R)
            # 对空格更加宽容，允许最多3个空格
            question_patterns = [
                # 标准数字编号 + 双引号题目：1. "Question text" (R)
                r'(\d+)\.\s{0,3}"(.*?)"\s{0,3}(?:\(R\))?\s{0,3}(?:\n|$)',
                # 带星号的数字编号 + 双引号题目：*1. "Question text" (R)
                r'\*?(\d+)\s{0,3}[.:]*\s{0,3}"(.*?)"\s{0,3}(?:\(R\))?\s{0,3}(?:\n|$)',
                # 项目符号 + 双引号题目：• "Question text" (R)
                r'•\s{0,3}"(.*?)"\s{0,3}(?:\(R\))?\s{0,3}(?:\n|$)',
                # 数字 + 空格 + 双引号题目：1 "Question text" (R)
                r'(\d+)\s{0,3}"(.*?)"\s{0,3}(?:\(R\))?\s{0,3}(?:\n|$)',
                # 前导星号 + 双引号题目：*"Question text" (R)
                r'\*\s{0,3}"(.*?)"\s{0,3}(?:\(R\))?\s{0,3}(?:\n|$)'
            ]
            
            for pattern in question_patterns:
                question_matches = re.findall(pattern, section_text, re.DOTALL)
                if question_matches:
                    for match in question_matches:
                        if len(match) == 2:
                            # Numbered question with quotes
                            q_num, q_text = match
                            # Clean up question number
                            q_num = re.sub(r'[^0-9]', '', q_num)
                            if q_num:
                                q_num = int(q_num)
                            else:
                                continue  # Skip if no valid number
                        else:
                            # Bullet point or other format question with quotes
                            if isinstance(match, tuple):
                                q_text = match[0]
                            else:
                                q_text = match
                            # Generate question number based on position
                            if dimension not in dimension_counter:
                                q_num = 1
                            else:
                                q_num = dimension_counter[dimension] + 1
                        
                        q_text = q_text.strip()
                        # 注意：由于正则表达式已经只捕获双引号内的文本，所以不需要再移除引号
                        # Remove any leading asterisks (if any)
                        q_text = q_text.lstrip('*').strip()
                        
                        # Check for reverse coding (more robust)
                        reverse_coded = False
                        if '(R)' in q_text:
                            reverse_coded = True
                            q_text = q_text.replace('(R)', '').strip()
                        elif '反向' in q_text:
                            reverse_coded = True
                            q_text = q_text.replace('反向', '').strip()
                        
                        # Generate question ID
                        if dimension not in dimension_counter:
                            dimension_counter[dimension] = 1
                        else:
                            dimension_counter[dimension] += 1
                        
                        # Create a short dimension code for question ID
                        dim_code = ''.join([word[0].upper() for word in dimension.split() if word])[:3]
                        question_id = f"{dim_code}_{dimension_counter[dimension]}"
                        
                        # Skip empty questions and non-question lines
                        non_question_markers = ['Items:', 'Question', 'Coding:', 'Scaling:', 'Scoring Key:', '或']
                        if q_text and len(q_text) > 3 and not any(marker in q_text for marker in non_question_markers):
                            questions.append({
                                "question_id": question_id,
                                "dimension": dimension,
                                "stem": q_text,
                                "coding": coding,
                                "reverse_coded": reverse_coded,
                                "score_range": score_range
                            })
                            # Print debug info
                            print(f"Parsed question: {question_id} - {q_text} (R: {reverse_coded})")
        
        # If no questions found, try alternative parsing approach
        if not questions:
            # Try parsing line by line with improved logic
            current_dimension = "Unknown"
            current_coding = "1-5 Likert scale"
            dimension_counter = {}
            
            for line in full_text:
                # Check for dimension headers
                is_dimension = False
                stripped_line = line.strip()
                
                if line.endswith(':') and not any(prefix in line for prefix in ['Items:', 'Question', 'Coding:', 'Scaling:', 'Scoring Key:']):
                    is_dimension = True
                else:
                    # 使用known_dimensions检查已知维度
                    if len(stripped_line) < 50:
                        for known_dim in known_dimensions:
                            if known_dim.lower() in stripped_line.lower():
                                if not any(prefix in line for prefix in ['Items:', 'Question', 'Coding:', 'Scaling:', 'Scoring Key:']):
                                    is_dimension = True
                                    break
                
                if is_dimension:
                    current_dimension = line.rstrip(':').strip()
                    current_dimension = re.sub(r'\s*(scale|Scale)$', '', current_dimension, flags=re.IGNORECASE)
                    if current_dimension not in dimension_counter:
                        dimension_counter[current_dimension] = 0
                
                # Check for coding information
                elif line.startswith('Coding:'):
                    current_coding = line[7:].strip()
                elif line.startswith('Scoring Key:'):
                    current_coding = line[13:].strip()
                
                # Check for questions
                elif re.match(r'^(\d+)\.|^•|^\*', line):
                    # Extract question text
                    if re.match(r'^(\d+)\.', line):
                        q_match = re.match(r'^(\d+)\.\s*(.*)$', line)
                        if q_match:
                            q_text = q_match.group(2).strip()
                    elif re.match(r'^•', line):
                        # Bullet point question
                        q_text = line[1:].strip()
                    elif re.match(r'^\*', line):
                        # Asterisk question
                        q_text = line[1:].strip()
                    else:
                        # Other format question
                        q_text = line.strip()
                    
                    # Remove quotes if present
                    if q_text.startswith('"') and q_text.endswith('"'):
                        q_text = q_text[1:-1]
                    
                    # Remove any leading asterisks
                    q_text = q_text.lstrip('*').strip()
                    
                    # Check for reverse coding
                    reverse_coded = '(R)' in q_text
                    if reverse_coded:
                        q_text = q_text.replace('(R)', '').strip()
                    elif '反向' in q_text:
                        reverse_coded = True
                        q_text = q_text.replace('反向', '').strip()
                    
                    # Generate question number
                    dimension_counter[current_dimension] += 1
                    q_num = dimension_counter[current_dimension]
                    
                    # Create a short dimension code for question ID
                    dim_code = ''.join([word[0].upper() for word in current_dimension.split() if word])[:3]
                    question_id = f"{dim_code}_{q_num}"
                    
                    # Determine score range
                    score_range = (1, 5)  # Default
                    if '7' in current_coding:
                        score_range = (1, 7)
                    
                    # Skip empty questions
                    if q_text:
                        questions.append({
                            "question_id": question_id,
                            "dimension": current_dimension,
                            "stem": q_text,
                            "coding": current_coding,
                            "reverse_coded": reverse_coded,
                            "score_range": score_range
                        })
        
        # If still no questions found, try LLM-based parsing
        if not questions:
            messagebox.showinfo("提示", "尝试使用大模型解析复杂问卷结构...")
            questions = parse_questionnaire_with_llm(document_text, token_limit)
        
        if not questions:
            messagebox.showerror("错误", "Word文件中未找到有效题目")
            return None
        
        print(f"\n=== 解析结果统计 ===")
        print(f"Successfully parsed {len(questions)} questions from Word file")
        
        # 按维度统计题目数量
        from collections import defaultdict
        dimension_question_count = defaultdict(int)
        for q in questions:
            dimension_question_count[q['dimension']] += 1
        
        print("\n各维度题目数量:")
        for dim, count in sorted(dimension_question_count.items()):
            print(f"  {dim}: {count} 题")
        
        print("\n前5道题的维度信息:")
        for i, q in enumerate(questions[:5], 1):
            print(f"  {i}. [{q['question_id']}] 维度: {q['dimension']}")
        
        return questions
        
    except Exception as e:
        messagebox.showerror("错误", f"解析Word问卷文件失败: {str(e)}")
        print(f"Error parsing Word file: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

def parse_questionnaire_with_llm(document_text, token_limit=4000):
    """Parse questionnaire using LLM for complex structures"""
    try:
        # Prepare prompt for LLM
        prompt = f"""You are an expert in questionnaire analysis. Please parse the following questionnaire text and extract:
1. Dimensions (questionnaire sections)
2. Questions within each dimension
3. Coding standards (scoring scales)
4. Reverse-coded items (marked with (R))

Return the result as a JSON array of objects, where each object has:
- question_id: Unique identifier (e.g., "EA_1" for Emotional Abuse question 1)
- dimension: Dimension name
- stem: Question text (without (R) marker)
- coding: Scoring standard
- reverse_coded: Boolean indicating if reverse-coded
- score_range: Tuple of (min, max) score values

Questionnaire text:
{document_text}

Please return only the JSON array, no other text."""
        
        # Call LLM with token limit
        response = call_llm(prompt, max_tokens=token_limit)
        
        # Parse JSON response
        import json
        # Extract JSON from response
        json_match = re.search(r'\[\s*\{[\s\S]*\}\s*\]', response)
        if not json_match:
            print("LLM response does not contain valid JSON")
            return None
        
        json_str = json_match.group(0)
        questions = json.loads(json_str)
        
        # Validate and format the result
        formatted_questions = []
        for q in questions:
            # Ensure required fields are present
            if all(key in q for key in ['question_id', 'dimension', 'stem', 'coding']):
                # Set default values for optional fields
                reverse_coded = q.get('reverse_coded', False)
                score_range = q.get('score_range', [1, 5])
                
                formatted_questions.append({
                    "question_id": q['question_id'],
                    "dimension": q['dimension'],
                    "stem": q['stem'],
                    "coding": q['coding'],
                    "reverse_coded": reverse_coded,
                    "score_range": tuple(score_range)
                })
        
        return formatted_questions if formatted_questions else None
        
    except Exception as e:
        print(f"Error parsing questionnaire with LLM: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

# ---------------- Main Process ----------------
def main():
    # 声明全局变量
    global FATAL_API_ERROR, DASHSCOPE_API_KEY, BASE_URL, MODEL_NAME, MAX_CONSECUTIVE_SAME_DIM
    
    # 清理所有可能的tkinter窗口，确保干净的状态
    try:
        import tkinter as tk
        # 尝试获取并销毁所有顶层窗口
        if tk._default_root:
            # 隐藏默认根窗口
            tk._default_root.withdraw()
            # 销毁所有子窗口
            for widget in tk._default_root.winfo_children():
                try:
                    widget.destroy()
                except:
                    pass
    except Exception as e:
        print(f"清理窗口时出现错误: {e}")
        pass
    
    # 重置全局错误标志
    FATAL_API_ERROR = False
    # 1. Show settings GUI
    settings = show_settings_gui()
    
    # Check if GUI was cancelled
    if not settings['questionnaire_file'] or not settings['subject_file'] or not settings['output_dir']:
        return
    
    # Update global API settings
    DASHSCOPE_API_KEY = settings['api_key']
    BASE_URL = settings['base_url']
    MODEL_NAME = settings['model_name']
    MAX_CONSECUTIVE_SAME_DIM = settings['max_consecutive']
    MAX_TOKENS = settings['max_tokens']
    
    # Re-initialize API client with new settings
    global client
    client = OpenAI(
        api_key=DASHSCOPE_API_KEY,
        base_url=BASE_URL,
    )
    
    # Extract settings
    questionnaire_file = settings['questionnaire_file']
    subject_file = settings['subject_file']
    output_dir = settings['output_dir']
    output_format = settings['output_format']
    random_order = settings['random_order']
    token_limit = settings['token_limit']
    max_tokens = settings['max_tokens']
    min_age = settings['min_age']
    max_age = settings['max_age']
    column_strategy = settings['column_strategy']
    
    print(f"Selected questionnaire file: {questionnaire_file}")
    print(f"Selected subject background file: {subject_file}")
    print(f"Selected output directory: {output_dir}")
    print(f"Random question order: {random_order}")
    print(f"Max consecutive same dimension: {MAX_CONSECUTIVE_SAME_DIM}")
    
    # 2. Parse questionnaire file
    questions = parse_questionnaire_file(questionnaire_file, token_limit)
    if not questions:
        messagebox.showerror("错误", "解析问卷文件失败，程序退出")
        return
    
    # 3. Create output directory
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    
    # 4. Load subject background
    subjects = load_subject_background(subject_file, output_dir, min_age, max_age)
    if not subjects:
        messagebox.showerror("错误", "未加载到有效被试，程序退出")
        return
    
    # 3. Iterate over subjects to generate responses
    all_results = []
    failed_records = []  # Record failed questions for later check
    
    # 创建进度条弹窗
    import tkinter as tk
    from tkinter import ttk
    import time
    
    progress_window = tk.Toplevel()
    progress_window.title("处理进度")
    progress_window.geometry("400x150")
    progress_window.resizable(False, False)
    
    # 确保窗口在最前面
    progress_window.attributes('-topmost', True)
    progress_window.lift()
    
    progress_label = tk.Label(progress_window, text="准备开始处理...", font=('Arial', 10))
    progress_label.pack(pady=20)
    
    progress_bar = ttk.Progressbar(progress_window, orient=tk.HORIZONTAL, length=350, mode='determinate')
    progress_bar.pack(pady=10)
    
    status_label = tk.Label(progress_window, text="", font=('Arial', 8), fg='gray')
    status_label.pack(pady=5)
    
    total_subjects = len(subjects)
    progress_bar['maximum'] = total_subjects
    
    # 先更新一次窗口让它显示出来
    progress_window.update()
    progress_window.update_idletasks()
    time.sleep(0.1)  # 短暂延迟确保窗口完全渲染
    
    # 添加标志来跟踪程序是否正常完成
    completed_successfully = False
    
    try:
        for i, subject in enumerate(subjects, 1):
            # Check fatal error: stop processing new subjects
            if FATAL_API_ERROR:
                break
            
            # 更新进度条
            progress_label.config(text=f"处理被试 {i}/{total_subjects}")
            status_label.config(text=f"正在处理被试 {subject['subject_id']} ({subject['性别']}, {subject['年龄']}岁)")
            progress_bar['value'] = i
            # 使用update_idletasks更轻量，避免卡顿
            progress_window.update_idletasks()
            
            print(f"\nProcessing subject {subject['subject_id']} ({subject['性别']}, {subject['年龄']} years old)...")
            subject_responses = []
            
            # Get question order based on settings
            if random_order:
                random_question_list = get_random_questions(questions)
                print(f"  Generated random question order (total {len(random_question_list)} questions)")
            else:
                # Use parsed question order (no randomization)
                random_question_list = questions  # 使用解析得到的问题列表
                print(f"  Using parsed question order (total {len(random_question_list)} questions)")
            
            # 并发处理所有问题以提高效率
            print(f"  开始并发处理 {len(random_question_list)} 个问题...")
            
            # 准备API设置
            api_settings = {
                'max_tokens': max_tokens
            }
            
            # 为每个问题添加随机序号
            for idx, question in enumerate(random_question_list, start=1):
                question['random_index'] = idx
            
            # 使用线程池并发处理
            max_workers = min(5, len(random_question_list))  # 限制并发数，避免API限制
            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                # 准备任务参数
                task_args = [(subject, question, column_strategy, api_settings) 
                            for question in random_question_list]
                
                # 提交所有任务
                future_to_question = {executor.submit(process_single_question, args): args 
                                    for args in task_args}
                
                # 处理完成的任务
                completed_count = 0
                for future in concurrent.futures.as_completed(future_to_question):
                    completed_count += 1
                    args = future_to_question[future]
                    question = args[1]
                    
                    # 检查致命错误
                    if FATAL_API_ERROR:
                        print(f"  检测到致命API错误，停止处理剩余问题")
                        break
                    
                    try:
                        response_record, error_record = future.result()
                        subject_responses.append(response_record)
                        
                        if error_record:
                            failed_records.append(error_record)
                        
                        print(f"  已完成 {completed_count}/{len(random_question_list)}: {question['question_id']} (状态: {response_record['作答状态']})")
                        
                    except Exception as e:
                        print(f"  处理问题 {question['question_id']} 时发生异常: {str(e)}")
                        
                        # 添加失败记录
                        failed_response = {
                            "被试ID": subject['subject_id'],
                            "性别": subject['性别'],
                            "年龄": subject['年龄'],
                            "随机题目序号": question.get('random_index', 0),
                            "原始题目ID": question['question_id'],
                            "维度": question['dimension'],
                            "题目内容（英文）": question['stem'],
                            "计分标准（英文）": question['coding'],
                            "是否反向计分": question['reverse_coded'],
                            "原始响应（英文）": f"PROCESSING_ERROR: {str(e)}",
                            "提取分数": None,
                            "最终得分": None,
                            "回答理由（英文）": "Processing error",
                            "作答状态": "失败"
                        }
                        
                        # 添加所有其他背景文件字段
                        for key, value in subject.items():
                            if key not in ['subject_id', '性别', '年龄']:
                                failed_response[key] = value
                        
                        subject_responses.append(failed_response)
                        failed_records.append({
                            "被试ID": subject['subject_id'],
                            "题目ID": question['question_id'],
                            "错误原因": str(e)
                        })
            
            # Calculate dimension scores for the subject
            scale_scores = calculate_scale_scores(subject_responses)
            # Merge dimension scores into each response
            for resp in subject_responses:
                resp.update(scale_scores)
            # Add to total results
            all_results.extend(subject_responses)
        
        # 所有被试处理完成，设置标志
        completed_successfully = True
    
    except KeyboardInterrupt:
        print("\n🔴 Program interrupted by user (Ctrl+C)")
    finally:
        # 关闭进度条弹窗
        try:
            progress_window.destroy()
        except:
            pass
        
        # 显示成功或错误弹窗
        from tkinter import messagebox
        
        if FATAL_API_ERROR:
            print(f"\n🔴 Program terminated due to fatal API error: {FATAL_ERROR_MSG}")
            print("🔴 Please resolve the API issue (e.g., recharge Alibaba Cloud account) and restart the program.")
            # 保存中断结果
            save_current_results(all_results, failed_records, out_dir, output_format, is_final=False)
            messagebox.showerror("错误", f"程序因API错误终止:\n{FATAL_ERROR_MSG}\n\n请解决API问题（例如，为阿里云账户充值）并重新启动程序。")
        elif not all_results:
            messagebox.showerror("错误", "未生成任何结果，请检查输入文件和设置。")
        elif not completed_successfully:
            # 程序被中断或出错
            print("\n🔴 Program did not complete successfully")
            save_current_results(all_results, failed_records, out_dir, output_format, is_final=False)
            messagebox.showwarning("警告", "程序未完全完成，已保存部分结果。")
        else:
            # 程序正常完成
            print("\n✅ Program exited safely (all current results saved)")
            # 保存最终结果
            save_current_results(all_results, failed_records, out_dir, output_format, is_final=True)
            output_file = out_dir / f"EasyPsych_Results.{output_format}"
            
            # 构建成功信息
            result_text = f"程序运行完成！\n\n已处理 {len(subjects)} 个被试\n已生成 {len(all_results)} 条结果\n\n结果文件保存位置:\n{output_file}"
            
            # 使用messagebox.showinfo显示信息，然后用askyesno询问下一步
            messagebox.showinfo("处理完成", result_text)
            
            # 询问用户是否要返回设置界面
            answer = messagebox.askyesno("选择下一步", "是否要返回设置界面重新测试？\n\n是 - 返回设置\n否 - 退出程序")
            
            if answer:
                # 用户选择返回设置
                # 返回True表示需要重新运行
                return True
            else:
                # 用户选择退出程序
                return False
    
    # 如果程序没有正常完成或者用户不返回设置，返回False
    return False

if __name__ == "__main__":
    # Create necessary directories for application packaging
    app_dir = Path(os.path.dirname(os.path.abspath(__file__)))
    icon_dir = app_dir / "icons"
    icon_dir.mkdir(exist_ok=True)
    
    # Create a placeholder for icon file
    icon_placeholder = icon_dir / "app_icon.png"
    if not icon_placeholder.exists():
        with open(icon_placeholder, 'w') as f:
            f.write("# App icon placeholder\n# Replace this file with your actual app icon (PNG format)")
    
    print(f"Created icon directory: {icon_dir}")
    print(f"Icon placeholder created at: {icon_placeholder}")
    
    # PyInstaller is only needed for building, not for running
    # 只在构建时需要，运行时不需要导入
    # Note: PyInstaller import is commented out to avoid Pylance warnings
    # It will be imported dynamically only when needed for building
    # def check_pyinstaller():
    #     try:
    #         # 尝试导入，但即使失败也不影响运行
    #         import PyInstaller
    #         print("PyInstaller is installed (for building)")
    #     except ImportError:
    #         # 静默处理，不打印任何信息
    #         pass
    
    # Only check PyInstaller if needed
    # check_pyinstaller()  # Uncomment if you want to check PyInstaller installation

    # 若为本地调试模式，生成一个小的受试者 Excel 供脚本读取（避免依赖外部文件）
    if 'DEBUG_MODE' in globals() and DEBUG_MODE:
        test_file = Path(OUTPUT_DIR) / "debug_test_subjects.xlsx"
        if not test_file.exists():
            df_test = pd.DataFrame([
                {
                    '性别': '女', '年龄': 30, '最高教育水平': '学士及以上学位',
                    '职业': '专业技术类', '行业': '专业及相关服务', '家庭年总收入': '$50,000–$74,999'
                },
                {
                    '性别': '男', '年龄': 45, '最高教育水平': '高中毕业',
                    '职业': '服务行业', '行业': '个人服务', '家庭年总收入': '$25,000–$49,999'
                }
            ])
            df_test.to_excel(test_file, index=False, engine='openpyxl')
            print(f"⚙️ DEBUG: 生成测试受试者文件 -> {test_file}")
        # 覆盖全局 SUBJECT_BACKGROUND_FILE 指向测试文件
        SUBJECT_BACKGROUND_FILE = str(test_file)

    # Run main process in a loop until user chooses to exit
    while True:
        try:
            should_restart = main()
            if not should_restart:
                break
        except KeyboardInterrupt:
            print("\n程序被用户中断")
            break
        except Exception as e:
            print(f"\n程序发生错误: {e}")
            import traceback
            traceback.print_exc()
            break