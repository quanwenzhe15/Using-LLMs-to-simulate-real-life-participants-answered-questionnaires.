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
#import random
import time
import pandas as pd
from pathlib import Path
from openai import OpenAI
from datetime import datetime
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

# ---------------- Core Configuration (Adjust as Needed) ----------------
# API Configuration (Alibaba Cloud Qwen)
DASHSCOPE_API_KEY = "sk-51b0406a9d884aa0aa99627d50a61329"  # Your API key
BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"  # Beijing region (no modification needed)
MODEL_NAME = "qwen-plus"  # Fixed model name

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

# ---------------- Questionnaire Items (Target Dimensions, English Version) ----------------
QUESTIONS = [
    # 1. 情感虐待（Emotional Abuse）- 5题，5点计分 1=Never true;5=Very often true，无反向，分数越高虐待越严重
    {
        "question_id": "EA_1",
        "dimension": "情感虐待",
        "stem": "People in my family called me things like “stupid,” “lazy,” or “ugly.” (When I was growing up)",
        "coding": "1=Never true; 2=Rarely true; 3=Sometimes true; 4=Often true; 5=Very often true",
        "reverse_coded": False,
        "score_range": (1, 5)
    },
    {
        "question_id": "EA_2",
        "dimension": "情感虐待",
        "stem": "I thought that my parents wished I had never been born. (When I was growing up)",
        "coding": "1=Never true; 2=Rarely true; 3=Sometimes true; 4=Often true; 5=Very often true",
        "reverse_coded": False,
        "score_range": (1, 5)
    },
    {
        "question_id": "EA_3",
        "dimension": "情感虐待",
        "stem": "People in my family said hurtful or insulting things to me. (When I was growing up)",
        "coding": "1=Never true; 2=Rarely true; 3=Sometimes true; 4=Often true; 5=Very often true",
        "reverse_coded": False,
        "score_range": (1, 5)
    },
    {
        "question_id": "EA_4",
        "dimension": "情感虐待",
        "stem": "I felt that someone in my family hated me. (When I was growing up)",
        "coding": "1=Never true; 2=Rarely true; 3=Sometimes true; 4=Often true; 5=Very often true",
        "reverse_coded": False,
        "score_range": (1, 5)
    },
    {
        "question_id": "EA_5",
        "dimension": "情感虐待",
        "stem": "I believe that I was emotionally abused. (When I was growing up)",
        "coding": "1=Never true; 2=Rarely true; 3=Sometimes true; 4=Often true; 5=Very often true",
        "reverse_coded": False,
        "score_range": (1, 5)
    },
    # 2. 情感忽视（Emotional Neglect）- 5题，5点计分 1=Never true;5=Very often true，全反向，分数越高忽视越严重
    {
        "question_id": "EN_1",
        "dimension": "情感忽视",
        "stem": "There was someone in my family who helped me feel that I was important or special. (When I was growing up)",
        "coding": "1=Never true; 2=Rarely true; 3=Sometimes true; 4=Often true; 5=Very often true",
        "reverse_coded": True,
        "score_range": (1, 5)
    },
    {
        "question_id": "EN_2",
        "dimension": "情感忽视",
        "stem": "I felt loved. (When I was growing up)",
        "coding": "1=Never true; 2=Rarely true; 3=Sometimes true; 4=Often true; 5=Very often true",
        "reverse_coded": True,
        "score_range": (1, 5)
    },
    {
        "question_id": "EN_3",
        "dimension": "情感忽视",
        "stem": "People in my family looked out for each other. (When I was growing up)",
        "coding": "1=Never true; 2=Rarely true; 3=Sometimes true; 4=Often true; 5=Very often true",
        "reverse_coded": True,
        "score_range": (1, 5)
    },
    {
        "question_id": "EN_4",
        "dimension": "情感忽视",
        "stem": "People in my family felt close to each other. (When I was growing up)",
        "coding": "1=Never true; 2=Rarely true; 3=Sometimes true; 4=Often true; 5=Very often true",
        "reverse_coded": True,
        "score_range": (1, 5)
    },
    {
        "question_id": "EN_5",
        "dimension": "情感忽视",
        "stem": "My family was a source of strength and support. (When I was growing up)",
        "coding": "1=Never true; 2=Rarely true; 3=Sometimes true; 4=Often true; 5=Very often true",
        "reverse_coded": True,
        "score_range": (1, 5)
    },
    # 3. 主管支持（Supervisory Support Scale）- 9题，5点李克特，无反向计分，分数越高支持度越高（原有正确，保留）
    {
        "question_id": "SS_1",
        "dimension": "主管支持",
        "stem": "My supervisor takes the time to learn about my career goals and aspirations",
        "coding": "1=strongly agree; 2=agree to some extent; 3=uncertain; 4=disagree to some extent; 5=strongly disagree",
        "reverse_coded": False,
        "score_range": (1, 5)
    },
    {
        "question_id": "SS_2",
        "dimension": "主管支持",
        "stem": "My supervisor cares about whether or not I achieve my goals",
        "coding": "1=strongly agree; 2=agree to some extent; 3=uncertain; 4=disagree to some extent; 5=strongly disagree",
        "reverse_coded": False,
        "score_range": (1, 5)
    },
    {
        "question_id": "SS_3",
        "dimension": "主管支持",
        "stem": "My supervisor keeps me informed about different career opportunities for me in the organization",
        "coding": "1=strongly agree; 2=agree to some extent; 3=uncertain; 4=disagree to some extent; 5=strongly disagree",
        "reverse_coded": False,
        "score_range": (1, 5)
    },
    {
        "question_id": "SS_4",
        "dimension": "主管支持",
        "stem": "My supervisor makes sure I get the credit when I accomplish something substantial on the job",
        "coding": "1=strongly agree; 2=agree to some extent; 3=uncertain; 4=disagree to some extent; 5=strongly disagree",
        "reverse_coded": False,
        "score_range": (1, 5)
    },
    {
        "question_id": "SS_5",
        "dimension": "主管支持",
        "stem": "My supervisor gives me helpful feedback about my performance",
        "coding": "1=strongly agree; 2=agree to some extent; 3=uncertain; 4=disagree to some extent; 5=strongly disagree",
        "reverse_coded": False,
        "score_range": (1, 5)
    },
    {
        "question_id": "SS_6",
        "dimension": "主管支持",
        "stem": "My supervisor gives me helpful advice about improving my performance when I need it",
        "coding": "1=strongly agree; 2=agree to some extent; 3=uncertain; 4=disagree to some extent; 5=strongly disagree",
        "reverse_coded": False,
        "score_range": (1, 5)
    },
    {
        "question_id": "SS_7",
        "dimension": "主管支持",
        "stem": "My supervisor supports my attempts to acquire additional training or education to further my career",
        "coding": "1=strongly agree; 2=agree to some extent; 3=uncertain; 4=disagree to some extent; 5=strongly disagree",
        "reverse_coded": False,
        "score_range": (1, 5)
    },
    {
        "question_id": "SS_8",
        "dimension": "主管支持",
        "stem": "My supervisor provides assignments that give me the opportunity to develop and strengthen new skills",
        "coding": "1=strongly agree; 2=agree to some extent; 3=uncertain; 4=disagree to some extent; 5=strongly disagree",
        "reverse_coded": False,
        "score_range": (1, 5)
    },
    {
        "question_id": "SS_9",
        "dimension": "主管支持",
        "stem": "My supervisor assigns me special projects that increase my visibility in the organization",
        "coding": "1=strongly agree; 2=agree to some extent; 3=uncertain; 4=disagree to some extent; 5=strongly disagree",
        "reverse_coded": False,
        "score_range": (1, 5)
    },
    # 4. 个人掌控（Personal Mastery）- 4题，7点李克特，反向计分，分数越高掌控感越强（原有正确，保留）
    {
        "question_id": "PM_1",
        "dimension": "个人掌控",
        "stem": "I can do just about anything I really set my mind to.",
        "coding": "1=Strongly agree; 2=Somewhat agree; 3=A little agree; 4=Don't know; 5=A little disagree; 6=Somewhat disagree; 7=Strongly disagree",
        "reverse_coded": True,
        "score_range": (1, 7)
    },
    {
        "question_id": "PM_2",
        "dimension": "个人掌控",
        "stem": "When I really want to do something, I usually find a way to succeed at it.",
        "coding": "1=Strongly agree; 2=Somewhat agree; 3=A little agree; 4=Don't know; 5=A little disagree; 6=Somewhat disagree; 7=Strongly disagree",
        "reverse_coded": True,
        "score_range": (1, 7)
    },
    {
        "question_id": "PM_3",
        "dimension": "个人掌控",
        "stem": "Whether or not I am able to get what I want is in my own hands.",
        "coding": "1=Strongly agree; 2=Somewhat agree; 3=A little agree; 4=Don't know; 5=A little disagree; 6=Somewhat disagree; 7=Strongly disagree",
        "reverse_coded": True,
        "score_range": (1, 7)
    },
    {
        "question_id": "PM_4",
        "dimension": "个人掌控",
        "stem": "What happens to me in the future mostly depends on me.",
        "coding": "1=Strongly agree; 2=Somewhat agree; 3=A little agree; 4=Don't know; 5=A little disagree; 6=Somewhat disagree; 7=Strongly disagree",
        "reverse_coded": True,
        "score_range": (1, 7)
    },
    # 5. 感知约束（Perceived Constraints）- 8题，7点李克特，无反向计分，分数越高约束感越强（原有正确，保留）
    {
        "question_id": "PC_1",
        "dimension": "感知约束",
        "stem": "There is little I can do to change the important things in my life.",
        "coding": "1=Strongly agree; 2=Somewhat agree; 3=A little agree; 4=Don't know; 5=A little disagree; 6=Somewhat disagree; 7=Strongly disagree",
        "reverse_coded": False,
        "score_range": (1, 7)
    },
    {
        "question_id": "PC_2",
        "dimension": "感知约束",
        "stem": "I often feel helpless in dealing with the problems of life.",
        "coding": "1=Strongly agree; 2=Somewhat agree; 3=A little agree; 4=Don't know; 5=A little disagree; 6=Somewhat disagree; 7=Strongly disagree",
        "reverse_coded": False,
        "score_range": (1, 7)
    },
    {
        "question_id": "PC_3",
        "dimension": "感知约束",
        "stem": "Other people determine most of what I can and cannot do.",
        "coding": "1=Strongly agree; 2=Somewhat agree; 3=A little agree; 4=Don't know; 5=A little disagree; 6=Somewhat disagree; 7=Strongly disagree",
        "reverse_coded": False,
        "score_range": (1, 7)
    },
    {
        "question_id": "PC_4",
        "dimension": "感知约束",
        "stem": "What happens in my life is often beyond my control.",
        "coding": "1=Strongly agree; 2=Somewhat agree; 3=A little agree; 4=Don't know; 5=A little disagree; 6=Somewhat disagree; 7=Strongly disagree",
        "reverse_coded": False,
        "score_range": (1, 7)
    },
    {
        "question_id": "PC_5",
        "dimension": "感知约束",
        "stem": "There are many things that interfere with what I want to do.",
        "coding": "1=Strongly agree; 2=Somewhat agree; 3=A little agree; 4=Don't know; 5=A little disagree; 6=Somewhat disagree; 7=Strongly disagree",
        "reverse_coded": False,
        "score_range": (1, 7)
    },
    {
        "question_id": "PC_6",
        "dimension": "感知约束",
        "stem": "I have little control over the things that happen to me.",
        "coding": "1=Strongly agree; 2=Somewhat agree; 3=A little agree; 4=Don't know; 5=A little disagree; 6=Somewhat disagree; 7=Strongly disagree",
        "reverse_coded": False,
        "score_range": (1, 7)
    },
    {
        "question_id": "PC_7",
        "dimension": "感知约束",
        "stem": "There is really no way I can solve the problems I have.",
        "coding": "1=Strongly agree; 2=Somewhat agree; 3=A little agree; 4=Don't know; 5=A little disagree; 6=Somewhat disagree; 7=Strongly disagree",
        "reverse_coded": False,
        "score_range": (1, 7)
    },
    {
        "question_id": "PC_8",
        "dimension": "感知约束",
        "stem": "I sometimes feel I am being pushed around in my life.",
        "coding": "1=Strongly agree; 2=Somewhat agree; 3=A little agree; 4=Don't know; 5=A little disagree; 6=Somewhat disagree; 7=Strongly disagree",
        "reverse_coded": False,
        "score_range": (1, 7)
    },
    # 6. 工作不安全感（Job Insecurity Scale）- 4题，5点李克特 1=Strongly disagree;5=Strongly agree，第4题反向，分数越高不安全感越强（修正为量表版）
    {
        "question_id": "JI_1",
        "dimension": "工作不安全感",
        "stem": "Chances are, I will soon lose my job.",
        "coding": "1=Strongly disagree; 2=Disagree; 3=Neither agree nor disagree; 4=Agree; 5=Strongly agree",
        "reverse_coded": False,
        "score_range": (1, 5)
    },
    {
        "question_id": "JI_2",
        "dimension": "工作不安全感",
        "stem": "I feel insecure about the future of my job.",
        "coding": "1=Strongly disagree; 2=Disagree; 3=Neither agree nor disagree; 4=Agree; 5=Strongly agree",
        "reverse_coded": False,
        "score_range": (1, 5)
    },
    {
        "question_id": "JI_3",
        "dimension": "工作不安全感",
        "stem": "I think I might lose my job in the near future.",
        "coding": "1=Strongly disagree; 2=Disagree; 3=Neither agree nor disagree; 4=Agree; 5=Strongly agree",
        "reverse_coded": False,
        "score_range": (1, 5)
    },
    {
        "question_id": "JI_4",
        "dimension": "工作不安全感",
        "stem": "I am sure I can keep my job.",
        "coding": "1=Strongly disagree; 2=Disagree; 3=Neither agree nor disagree; 4=Agree; 5=Strongly agree",
        "reverse_coded": True,
        "score_range": (1, 5)
    }
]

# ---------------- Tool Functions ----------------
def load_subject_background(file_path):
    """Read subject background Excel, return standardized subject list"""
    print(f"Reading subject background file: {file_path}")
    try:
        df = pd.read_excel(file_path)
        required_cols = ['性别', '年龄', '最高教育水平', '职业', '行业', '家庭年总收入']
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            raise ValueError(f"Excel missing required columns: {', '.join(missing_cols)} (ensure header matches requirements)")
        
        # 1. 统一处理缺失值：把文本「缺失值」替换成NaN，方便后续过滤
        df = df.replace("缺失值", pd.NA)
        
        # 2. 年龄列清洗：转数值类型，过滤18-75岁的有效成年被试
        df['年龄'] = pd.to_numeric(df['年龄'], errors='coerce').astype('Int64')
        df = df[(df['年龄'] >= 18) & (df['年龄'] <= 75)]
        
        # 3. 文本列安全处理：先转字符串，再strip
        text_cols = ['性别', '最高教育水平', '职业', '行业']
        for col in text_cols:
            df[col] = df[col].fillna("不适用").astype(str).str.strip()
        
        # 4. 家庭年收入列特殊处理：数值转字符串，缺失值统一为"不适用"
        df['家庭年总收入'] = df['家庭年总收入'].apply(
            lambda x: f"{int(x)}" if pd.notna(x) and isinstance(x, (int, float)) else "不适用"
        )
        
        # 5. 过滤核心字段全空的行
        df = df.dropna(subset=['性别', '年龄', '最高教育水平'])
        
        # Convert to subject list
        subjects = []
        for idx, row in df.iterrows():
            subjects.append({
                "subject_id": int(row['被试ID']) if pd.notna(row['被试ID']) else idx + 1,
                "性别": row['性别'],
                "年龄": row['年龄'],
                "最高教育水平": row['最高教育水平'],
                "职业": row['职业'],
                "行业": row['行业'],
                "家庭年总收入": row['家庭年总收入']
            })
        
        print(f"Successfully loaded {len(subjects)} valid subjects (excluded nulls/invalid ages)")
        return subjects
    except Exception as e:
        print(f"Failed to read subject background: {str(e)}")
        import traceback
        traceback.print_exc()  
        return []

def generate_subject_prompt(subject, question):
    """Generate subject-specific prompt (English, adapted for American context)"""
    # 优化主管支持备注：根据职业是否为缺失/不适用判断
    supervisor_note = ""
    if "主管支持" in question['dimension']:
        if subject['职业'] in ["不适用", "拒绝回答", "不知道"]:
            supervisor_note = " (Note: If you don't have a supervisor or job, answer based on hypothetical work experience or common sense)"
        else:
            supervisor_note = f" (Note: Answer combined with your occupation as {subject['职业']} in {subject['行业']} industry)"
    
    # English prompt template
    prompt = f"""You are a real American citizen with the following personal background:
- Gender: {subject['性别']}
- Age: {subject['年龄']} years old
- Highest Education Level: {subject['最高教育水平']}
- Occupation: {subject['职业']}
- Industry: {subject['行业']}
- Annual Household Income: {subject['家庭年总收入']}
Fully embody this role, combine American cultural background, life experiences, and true feelings to answer the following questionnaire in the first person{supervisor_note}. Response requirements:
1. Strictly select a score based on the given coding standard (only enter a number between {question['score_range'][0]}-{question['score_range'][1]});
2. Add 1-2 sentences to explain the reason after the score. The reason should match your occupation, industry, income level and American social culture, avoiding emptiness;
3. Answer naturally and colloquially, like an ordinary American chatting—no formal writing or AI tone;
4. For work-related questions, answer based on your occupation, industry and career experience in the U.S.;
5. Do not reveal you are a simulated role, and never say phrases like "as an AI" or "according to the setting";
6. Only answer based on the current task, do not reference any previous responses.
Question: {question['stem']}
Coding Standard: {question['coding']}
Please answer directly without additional formatting."""
    return prompt

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
def call_llm(prompt):
    """Call Qwen API with retry mechanism, return raw response"""
    global FATAL_API_ERROR, FATAL_ERROR_MSG
    try:
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {"role": "user", "content": prompt}  # 补全你代码截断的messages部分
            ],
            max_tokens=MAX_TOKENS,
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
    
    # 1. 情感虐待（5题）
    ea_scores = dimension_groups.get('情感虐待', [])
    scale_scores['情感虐待_总分'] = sum(ea_scores) if len(ea_scores) == 5 else None
    scale_scores['情感虐待_平均分'] = round(sum(ea_scores)/len(ea_scores), 2) if len(ea_scores) == 5 else None
    
    # 2. 情感忽视（5题）
    en_scores = dimension_groups.get('情感忽视', [])
    scale_scores['情感忽视_总分'] = sum(en_scores) if len(en_scores) == 5 else None
    scale_scores['情感忽视_平均分'] = round(sum(en_scores)/len(en_scores), 2) if len(en_scores) == 5 else None
    
    # 3. 主管支持（9题）
    ss_scores = dimension_groups.get('主管支持', [])
    scale_scores['主管支持_总分'] = sum(ss_scores) if len(ss_scores) == 9 else None
    scale_scores['主管支持_平均分'] = round(sum(ss_scores)/len(ss_scores), 2) if len(ss_scores) == 9 else None
    
    # 4. 个人掌控（4题）
    pm_scores = dimension_groups.get('个人掌控', [])
    scale_scores['个人掌控_总分'] = sum(pm_scores) if len(pm_scores) == 4 else None
    scale_scores['个人掌控_平均分'] = round(sum(pm_scores)/len(pm_scores), 2) if len(pm_scores) == 4 else None
    
    # 5. 感知约束（8题）
    pc_scores = dimension_groups.get('感知约束', [])
    scale_scores['感知约束_总分'] = sum(pc_scores) if len(pc_scores) == 8 else None
    scale_scores['感知约束_平均分'] = round(sum(pc_scores)/len(pc_scores), 2) if len(pc_scores) == 8 else None
    
    # 6. 工作不安全感（4题，修正为量表版）
    ji_scores = dimension_groups.get('工作不安全感', [])
    scale_scores['工作不安全感_总分'] = sum(ji_scores) if len(ji_scores) == 4 else None
    scale_scores['工作不安全感_平均分'] = round(sum(ji_scores)/len(ji_scores), 2) if len(ji_scores) == 4 else None
    
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

#def get_random_questions(original_questions):
    """Generate random question order with constraint: no same dimension for 4 consecutive times"""
    while True:
        # Create a copy to avoid modifying original list
        random_questions = original_questions.copy()
        random.shuffle(random_questions)
        
        # Check if constraint is satisfied
        valid = True
        for i in range(len(random_questions) - MAX_CONSECUTIVE_SAME_DIM):
            # Get current dimension and next 3 dimensions (total 4 consecutive)
            current_dim = random_questions[i]['dimension']
            consecutive_dims = [random_questions[j]['dimension'] for j in range(i, i + MAX_CONSECUTIVE_SAME_DIM + 1)]
            
            # If all 4 are same dimension, invalid
            if all(dim == current_dim for dim in consecutive_dims):
                valid = False
                break
        
        if valid:
            return random_questions

def save_current_results(all_results, failed_records, out_dir):
    """Save current results immediately (even if process is stopped)"""
    if all_results:
        df_out = pd.DataFrame(all_results)
        # Adjust column order for readability
        # 修复：所有逗号改为英文半角，补全列名分隔符
        column_order = [
        "被试ID", "性别", "年龄", "教育水平",
        "职业", "行业", "家庭年总收入",
        "随机题目序号", "原始题目ID", "维度", "题目内容（英文）", "计分标准（英文）", "是否反向计分",
        "原始响应（英文）", "提取分数", "最终得分", "回答理由（英文）", "作答状态",
        "情感虐待_总分", "情感虐待_平均分", "情感忽视_总分", "情感忽视_平均分",
        "主管支持_总分", "主管支持_平均分", "个人掌控_总分", "个人掌控_平均分",
        "感知约束_总分", "感知约束_平均分", "工作不安全感_总分", "工作不安全感_平均分"
        ]
        # Ensure all columns exist
        for col in column_order:
            if col not in df_out.columns:
                df_out[col] = None
        df_out = df_out[column_order]
        
        # Generate filename with timestamp (mark as interrupted)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = out_dir / f"Interrupted_Results_{timestamp}.xlsx"
        
        # Save Excel
        df_out.to_excel(output_file, index=False, engine='openpyxl')
        print(f"\n Current results saved to: {output_file}")
        
        # Save failed records if any
        if failed_records:
            df_failed = pd.DataFrame(failed_records)
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
            error_file = out_dir / f"Fatal_Error_Info_{timestamp}.xlsx"
            error_info.to_excel(error_file, index=False, engine='openpyxl')
            print(f"✅ Fatal error info saved to: {error_file}")
    else:
        print("\n⚠️ No results to save (all_results is empty)")

# ---------------- Main Process ----------------
def main():
    global FATAL_API_ERROR
    # 1. Load subject background
    subjects = load_subject_background(SUBJECT_BACKGROUND_FILE)
    if not subjects:
        print("No valid subjects, program exited")
        return
    
    # 2. Create output directory
    out_dir = Path(OUTPUT_DIR)
    out_dir.mkdir(parents=True, exist_ok=True)
    
    # 3. Iterate over subjects to generate responses
    all_results = []
    failed_records = []  # Record failed questions for later check
    
    try:
        for subject in subjects:
            # Check fatal error: stop processing new subjects
            if FATAL_API_ERROR:
                break
            
            print(f"\nProcessing subject {subject['subject_id']} ({subject['性别']}, {subject['年龄']} years old)...")
            subject_responses = []
            
            # Get random question order (satisfy dimension constraint)
            #random_question_list = get_random_questions(QUESTIONS)
            #print(f"  Generated random question order (total {len(random_question_list)} questions)")
            # Use original question order (no randomization)
            random_question_list = QUESTIONS  # 直接使用原始QUESTIONS列表的顺序
            print(f"  Using original question order (total {len(random_question_list)} questions)")
            
            # Answer questions in random order
            for idx, question in enumerate(random_question_list, start=1):
                # Check fatal error: stop processing new questions for current subject
                if FATAL_API_ERROR:
                    break
                
                print(f"  Answering question {idx}/{len(random_question_list)}: {question['question_id']} (Dimension: {question['dimension']})")
                try:
                    # Generate prompt
                    prompt = generate_subject_prompt(subject, question)
                    # The following block is likely intended for exception handling, so wrap it in except
                    # Simulate API call and response parsing (replace with actual API call logic)
                    raw_resp = call_llm(prompt)
                    score, reason = parse_question_response(raw_resp, question)
                    subject_responses.append({
                        "被试ID": subject['subject_id'],
                        "性别": subject['性别'],
                        "职业": subject['职业'],
                        "行业": subject['行业'],
                        "家庭年总收入": subject['家庭年总收入'],
                        "年龄": subject['年龄'],
                        "教育水平": subject['最高教育水平'],
                        "随机题目序号": idx,
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
                    })
                except Exception as error_msg:
                    # Add to failed records
                    subject_responses.append({
                        "被试ID": subject['subject_id'],
                        "性别": subject['性别'],
                        "职业": subject['职业'],
                        "行业": subject['行业'],
                        "家庭年总收入": subject['家庭年总收入'],
                        "年龄": subject['年龄'],
                        "教育水平": subject['最高教育水平'],
                        "随机题目序号": idx,
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
                    })
                    failed_records.append({
                        "被试ID": subject['subject_id'],
                        "题目ID": question['question_id'],
                        "错误原因": str(error_msg)
                    })
            
            # Calculate dimension scores for the subject
            scale_scores = calculate_scale_scores(subject_responses)
            # Merge dimension scores into each response
            for resp in subject_responses:
                resp.update(scale_scores)
            # Add to total results
            all_results.extend(subject_responses)
    
    except KeyboardInterrupt:
        print("\n🔴 Program interrupted by user (Ctrl+C)")
    finally:
        # Save current results no matter why process stopped
        save_current_results(all_results, failed_records, out_dir)
        if FATAL_API_ERROR:
            print(f"\n🔴 Program terminated due to fatal API error: {FATAL_ERROR_MSG}")
            print("🔴 Please resolve the API issue (e.g., recharge Alibaba Cloud account) and restart the program.")
        print("\n✅ Program exited safely (all current results saved)")

if __name__ == "__main__":
    # Ensure required dependency 'tenacity' is available
    try:
        from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
    except ImportError:
        print("Installing required package 'tenacity'...")
        os.system("pip install tenacity")
        from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

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

    # Run main process
    main()