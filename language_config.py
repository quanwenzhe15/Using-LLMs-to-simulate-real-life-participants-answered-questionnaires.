# EasyPsych 多语言配置文件
import locale

def detect_system_language():
    """检测系统语言并返回对应的语言代码"""
    try:
        # 获取系统默认语言
        system_lang, _ = locale.getdefaultlocale()
        if system_lang:
            # 检查语言代码
            if 'zh' in system_lang.lower():
                return 'zh'
            elif 'en' in system_lang.lower():
                return 'en'
    except:
        pass
    
    # 默认返回英文
    return 'en'

LANGUAGE_CONFIG = {
    "zh": {
        # GUI界面文本
        "welcome_title": "问卷模拟系统 - 欢迎",
        "welcome_message": "欢迎使用EasyPsych问卷模拟系统",
        "api_settings": "API设置",
        "questionnaire_settings": "问卷设置",
        "file_selection": "文件选择",
        "start_processing": "开始运行",
        "edit_prompt_template": "编辑提示模板",
        "cancel": "取消",
        "save": "保存",
        
        # API设置标签页
        "api_key": "API密钥:",
        "base_url": "基础URL:",
        "model_name": "模型名称:",
        "max_tokens": "单次回答最大token数:",
        "column_strategy": "列策略:",
        
        # 问卷设置标签页
        "random_question_order": "启用随机题目顺序",
        "max_consecutive_same_dim": "同一维度最大连续出现数量:",
        "token_limit": "API分析最大token数:",
        "age_range": "被试年龄范围:",
        "min_age": "最小年龄:",
        "max_age": "最大年龄:",
        "custom_scoring_rules": "自定义计分规则:",
        
        # 文件选择标签页
        "questionnaire_file": "问卷文件:",
        "subject_file": "被试背景文件:",
        "output_dir": "输出结果路径:",
        "output_format": "输出格式:",
        "output_filename": "输出文件名:",
        "no_extension": "(无需扩展名)",
        "browse": "浏览",
        "select_questionnaire_file": "选择问卷文件",
        "select_subject_file": "选择被试背景文件",
        "select_output_dir": "选择输出结果路径",
        
        # 语言设置
        "language_settings": "语言设置 / Language:",
        "chinese": "中文",
        "english": "English",
        
        # 条款和条件
        "terms_agree": "我已阅读并同意上述条款和条件",
        "terms_must_agree": "您必须同意条款和条件才能继续使用本程序",
        "continue_button": "继续",
        "exit_button": "退出",
        
        # 问卷设置详细文本
        "api_token_limit": "API分析最大token数:",
        "tokens": "tokens",
        "max_tokens_per_response": "单次回答最大token数:",
        "subject_age_range": "被试年龄范围:",
        "scoring_rules": "自定义计分规则:",
        "scoring_rules_settings": "计分规则设置:",
        "edit_scoring_rules": "编辑计分规则",
        "new_column_strategy": "新列名处理策略:",
        
        # 进度条文本
        "progress_ready": "准备开始处理...",
        "progress_processing": "处理中...",
        
        # 错误消息
        "error_no_questionnaire": "未选择问卷文件",
        "error_no_subject": "未选择被试背景文件",
        "error_no_output": "未选择输出结果路径",
        "error_parsing_failed": "解析问卷文件失败",
        "error_no_valid_subjects": "未加载到有效被试",
        "error_api_fatal": "程序因API错误终止",
        "error_check_balance": "请检查API账户余额",
        "error_check_api_input": "请检查API密钥、基础URL和模型名称是否正确输入",
        "warning_incomplete": "程序未完全完成，已保存部分结果",
        
        # 成功消息
        "success_completed": "程序运行完成！",
        "success_subjects_processed": "已处理 {count} 个被试",
        "success_results_generated": "已生成 {count} 条结果",
        "success_file_saved": "结果文件保存位置:",
        "success": "处理完成",
        "next_step": "选择下一步",
        "return_to_settings": "是否要返回设置界面重新测试？\n\n是 - 返回设置\n否 - 退出程序",
        
        # 进度条
        "progress_title": "处理进度",
        "progress_ready": "准备开始处理...",
        "progress_processing": "处理被试 {current}/{total}",
        "progress_subject_info": "正在处理被试 {id} ({gender}, {age}岁)",
        
        # 错误和警告
        "error": "错误",
        "warning": "警告",
        
        # 提示模板
        "prompt_template": "提示模板:",
        
        # 列名映射（Excel/CSV文件）
        "question_id": "题目ID",
        "dimension": "题目所属维度", 
        "question_content": "题目内容",
        "scoring_standard": "计分标准"
    },
    
    "en": {
        # GUI界面文本
        "welcome_title": "Questionnaire Simulation System - Welcome",
        "welcome_message": "Welcome to EasyPsych Questionnaire Simulation System",
        "api_settings": "API Settings",
        "questionnaire_settings": "Questionnaire Settings", 
        "file_selection": "File Selection",
        "start_processing": "Start Processing",
        "edit_prompt_template": "Edit Prompt Template",
        "cancel": "Cancel",
        "save": "Save",
        
        # API设置标签页
        "api_key": "API Key:",
        "base_url": "Base URL:",
        "model_name": "Model Name:",
        "max_tokens": "Max Tokens per Response:",
        "column_strategy": "Column Strategy:",
        
        # 问卷设置标签页
        "random_question_order": "Enable Random Question Order",
        "max_consecutive_same_dim": "Max Consecutive Same Dimension:",
        "token_limit": "API Analysis Token Limit:",
        "age_range": "Subject Age Range:",
        "min_age": "Min Age:",
        "max_age": "Max Age:",
        "custom_scoring_rules": "Custom Scoring Rules:",
        
        # 文件选择标签页
        "questionnaire_file": "Questionnaire File:",
        "subject_file": "Subject Background File:",
        "output_dir": "Output Directory:",
        "output_format": "Output Format:",
        "output_filename": "Output Filename:",
        "no_extension": "(no extension needed)",
        "browse": "Browse",
        "select_questionnaire_file": "Select Questionnaire File",
        "select_subject_file": "Select Subject Background File",
        "select_output_dir": "Select Output Directory",
        
        # 语言设置
        "language_settings": "Language Settings:",
        "chinese": "中文",
        "english": "English",
        
        # 条款和条件
        "terms_agree": "I have read and agree to the above terms and conditions",
        "terms_must_agree": "You must agree to the terms and conditions to continue using this program",
        "continue_button": "Continue",
        "exit_button": "Exit",
        
        # 问卷设置详细文本
        "api_token_limit": "API Analysis Token Limit:",
        "tokens": "tokens",
        "max_tokens_per_response": "Max Tokens per Response:",
        "subject_age_range": "Subject Age Range:",
        "scoring_rules": "Custom Scoring Rules:",
        "scoring_rules_settings": "Scoring Rules Settings:",
        "edit_scoring_rules": "Edit Scoring Rules",
        "new_column_strategy": "New Column Name Strategy:",
        
        # 进度条文本
        "progress_ready": "Ready to start processing...",
        "progress_processing": "Processing...",
        
        # 错误消息
        "error_no_questionnaire": "No questionnaire file selected",
        "error_no_subject": "No subject background file selected",
        "error_no_output": "No output directory selected",
        "error_parsing_failed": "Failed to parse questionnaire file",
        "error_no_valid_subjects": "No valid subjects loaded",
        "error_api_fatal": "Program terminated due to API error",
        "error_check_balance": "Please check your API account balance",
        "error_check_api_input": "Please check if API key, base URL, and model name are correctly entered",
        "warning_incomplete": "Program did not complete fully, partial results saved",
        
        # 成功消息
        "success_completed": "Program completed successfully!",
        "success_subjects_processed": "Processed {count} subjects",
        "success_results_generated": "Generated {count} results",
        "success_file_saved": "Result file saved at:",
        "success": "Processing Completed",
        "next_step": "Next Step",
        "return_to_settings": "Do you want to return to the settings interface to retest?\n\nYes - Return to settings\nNo - Exit program",
        
        # 进度条
        "progress_title": "Processing Progress",
        "progress_ready": "Ready to start processing...",
        "progress_processing": "Processing subject {current}/{total}",
        "progress_subject_info": "Processing subject {id} ({gender}, {age} years old)",
        
        # 错误和警告
        "error": "Error",
        "warning": "Warning",
        
        # 提示模板
        "prompt_template": "Prompt Template:",
        
        # 列名映射（Excel/CSV文件）
        "question_id": "Question ID",
        "dimension": "Dimension",
        "question_content": "Question Content", 
        "scoring_standard": "Scoring Standard"
    }
}

# 当前语言设置（默认跟随系统语言）
CURRENT_LANGUAGE = detect_system_language()

def set_language(lang):
    """设置当前语言"""
    global CURRENT_LANGUAGE
    if lang in LANGUAGE_CONFIG:
        CURRENT_LANGUAGE = lang
    else:
        CURRENT_LANGUAGE = "zh"

def get_text(key, **kwargs):
    """获取本地化文本"""
    text = LANGUAGE_CONFIG[CURRENT_LANGUAGE].get(key, key)
    if kwargs:
        return text.format(**kwargs)
    return text

def get_column_names():
    """获取当前语言的列名映射"""
    return {
        "question_id": get_text("question_id"),
        "dimension": get_text("dimension"),
        "question_content": get_text("question_content"),
        "scoring_standard": get_text("scoring_standard")
    }