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
        # 欢迎界面
        "welcome_title": "EasyPsych - 心理学问卷模拟系统",
        "welcome_message": "欢迎使用 EasyPsych 心理学问卷模拟系统",
        "function_introduction": "功能介绍",
        "file_format_requirements": "文件格式要求",
        "usage_guide": "使用指南",
        "license_agreement": "许可证协议",
        "system_features_title": "本系统主要功能：",
        "feature_1": "支持多种问卷文件格式：Excel (.xlsx, .xls)、CSV (.csv) 和 Word (.docx)",
        "feature_2": "自动解析问卷结构，提取题目、维度、计分标准等信息",
        "feature_3": "支持随机题目顺序，可限制同一维度连续出现数量",
        "feature_4": "集成大模型 API，处理复杂问卷结构",
        "feature_5": "模拟被试回答，生成标准化结果文件",
        "feature_6": "支持反向计分题处理",
        "feature_7": "支持自定义计分规则，可根据需要修改计分标准",
        "file_format_requirements_title": "文件格式要求：",
        "questionnaire_file_title": "问卷文件：",
        "excel_file_requirement": "Excel 文件 (.xlsx, .xls)：必须包含以下列：题目ID、题目所属维度、题目内容、计分标准（或英文对应列）",
        "csv_file_requirement": "CSV 文件 (.csv)：必须包含以下列：题目ID、题目所属维度、题目内容、计分标准（或英文对应列）",
        "word_file_requirement": "Word 文件 (.docx)：按维度分节，维度标题以冒号结尾",
        "question_format_requirement": "题目格式要求：必须包含来回双引号，例如：4. \"People in my family felt close to each other.\" (R)",
        "supported_question_formats": "支持的题目格式：",
        "standard_numbering": "标准数字编号：1. \"Question text\" (R)",
        "with_asterisk": "带星号：*1. \"Question text\" (R)",
        "bullet_point": "项目符号：• \"Question text\" (R)",
        "with_space": "带空格：1 \"Question text\" (R)",
        "scoring_rule_format": "计分规则格式：Coding: 后连接的一句话，以句号结尾，例如：Coding: 1 Never true; 2 Rarely true; 3 Sometimes true; 4 Often true; 5 Very often true.",
        "reverse_scoring_markers": "支持反向计分标记：(R) 或 (反向)",
        "multiple_scoring_ranges": "支持多种计分范围：1-5、1-7、1-6等（自动识别）",
        "subject_background_file_title": "被试背景文件：",
        "supported_formats": "支持格式：Excel (.xlsx, .xls) 和 CSV (.csv)",
        "mandatory_columns": "强制要求的列：被试ID、年龄、性别（或英文对应列）",
        "other_columns": "其他列会被自动解析并加入到提示中",
        "null_value_handling": "空值会被填充为'不适用'，并生成缺失值报告",
        "high_missing_value_handling": "如果某列缺失值超过20%，会弹窗告知并让用户选择是否继续",
        "usage_guide_title": "使用指南：",
        "usage_step_1": "运行脚本后，在欢迎窗口中阅读并同意协议",
        "usage_step_2": "在设置窗口中配置 API 参数（如需要）",
        "usage_step_3": "选择是否启用随机题目顺序及连续维度限制",
        "usage_step_4": "选择问卷文件（Excel、CSV 或 Word 格式）",
        "usage_step_5": "选择被试背景文件（Excel 格式）",
        "usage_step_6": "选择输出结果路径",
        "usage_step_7": "点击'开始运行'按钮开始处理",
        "license_title": "GNU GENERAL PUBLIC LICENSE",
        "license_version": "Version 3, 29 June 2007",
        "license_paragraph_1": "本程序是自由软件：您可以根据自由软件基金会发布的 GNU 通用公共许可证条款",
        "license_paragraph_2": "（本许可证的第 3 版或您选择的任何更高版本）来重新分发和/或修改它。",
        "license_paragraph_3": "本程序的发布是希望它能有用，但没有任何担保；甚至没有对适销性或",
        "license_paragraph_4": "特定用途适用性的默示担保。有关详细信息，请参阅 GNU 通用公共许可证。",
        "license_paragraph_5": "您应该已经收到了一份 GNU 通用公共许可证的副本。如果没有，请参见",
        "license_website": "<https://www.gnu.org/licenses/>.",
        "missing_mandatory_columns": "背景文件缺少必要列",
        "no_valid_subjects_loaded": "未加载到有效被试，请检查后再次尝试",
        "warning_high_missing_values": "警告：高缺失值",
        "high_missing_value_columns": "以下列的缺失值超过20%：",
        "continue_running": "是否继续运行？",
        "warning_age_range": "警告：年龄范围检查",
        "age_out_of_range": "发现",
        "subjects": "个被试的年龄不在设定范围",
        "years": "岁",
        "invalid_ages": "无效年龄",
        "continue_with_filter": "是否继续运行（将自动过滤掉年龄无效的被试）？",
        "scoring_rules_title": "计分规则设置",
        "score_range_recognition_rule": "计分范围识别规则",
        "auto_recognize_score_range": "自动识别计分范围：根据计分标准中的数字确定",
        "example_score_range_1": "例如：包含 '7' 则为 1-7 点计分，包含 '6' 则为 1-6 点计分",
        "default_score_range": "默认为 1-5 点计分",
        "reverse_scoring_rule": "反向计分规则",
        "auto_recognize_reverse_markers": "自动识别反向计分标记：(R) 或 (反向)",
        "reverse_scoring_calculation": "反向计分计算：(最小值 + 最大值) - 原始分数",
        "example_reverse_scoring_1": "例如：5点计分中，原始分数为 1，则反向计分为 5",
        "example_reverse_scoring_2": "例如：7点计分中，原始分数为 2，则反向计分为 6",
        "dimension_score_calculation_rule": "维度分数计算规则",
        "dimension_score_sum": "维度分数 = 该维度下所有题目分数的总和",
        "missing_value_handling": "支持缺失值处理：仅一个缺失值时使用均值替换",
        "no_missing_values": "无缺失值时直接求和",
        "custom_rule_examples": "自定义规则示例",
        "example_modify_score_range": "示例1：修改计分范围识别",
        "force_5_point_scoring": "强制使用5点计分",
        "example_modify_reverse_calculation": "示例2：修改反向计分计算",
        "example_modify_dimension_calculation": "示例3：修改维度分数计算为平均值",
        "strategy_keep_original": "保持原样",
        "strategy_keep_original_desc": "保持中文字段名原样（推荐，AI能理解）",
        "strategy_auto_translate": "自动翻译",
        "strategy_auto_translate_desc": "尝试自动翻译为英文",
        "strategy_pinyin": "拼音转换",
        "strategy_pinyin_desc": "使用拼音作为英文标识",
        "strategy_custom_map": "自定义映射",
        "strategy_custom_map_desc": "在编辑提示模板时手动指定英文名",
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
        "pause_button": "暂停",
        "resume_button": "继续",
        "cancel_button": "取消",
        
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
        # 欢迎界面
        "welcome_title": "EasyPsych - Psychology Questionnaire Simulation System",
        "welcome_message": "Welcome to EasyPsych Psychology Questionnaire Simulation System",
        "function_introduction": "Function Introduction",
        "file_format_requirements": "File Format Requirements",
        "usage_guide": "Usage Guide",
        "license_agreement": "License Agreement",
        "system_features_title": "System Features:",
        "feature_1": "Supports multiple questionnaire file formats: Excel (.xlsx, .xls), CSV (.csv), and Word (.docx)",
        "feature_2": "Automatically parses questionnaire structure, extracts questions, dimensions, scoring criteria, etc.",
        "feature_3": "Supports random question order, can limit the number of consecutive same dimensions",
        "feature_4": "Integrates large model API to handle complex questionnaire structures",
        "feature_5": "Simulates subject responses, generates standardized result files",
        "feature_6": "Supports reverse scoring question processing",
        "feature_7": "Supports custom scoring rules, can modify scoring criteria as needed",
        "file_format_requirements_title": "File Format Requirements:",
        "questionnaire_file_title": "Questionnaire File:",
        "excel_file_requirement": "Excel file (.xlsx, .xls): Must include the following columns: 题目ID/Question ID, 题目所属维度/Dimension, 题目内容/Question Content, 计分标准/Scoring Criteria",
        "csv_file_requirement": "CSV file (.csv): Must include the following columns: 题目ID/Question ID, 题目所属维度/Dimension, 题目内容/Question Content, 计分标准/Scoring Criteria",
        "word_file_requirement": "Word file (.docx): Organized by dimensions, dimension titles ending with colon",
        "question_format_requirement": "Question format requirement: Must include double quotes, e.g.: 4. \"People in my family felt close to each other.\" (R)",
        "supported_question_formats": "Supported question formats:",
        "standard_numbering": "Standard numbering: 1. \"Question text\" (R)",
        "with_asterisk": "With asterisk: *1. \"Question text\" (R)",
        "bullet_point": "Bullet point: • \"Question text\" (R)",
        "with_space": "With space: 1 \"Question text\" (R)",
        "scoring_rule_format": "Scoring rule format: Coding: followed by a sentence ending with period, e.g.: Coding: 1 Never true; 2 Rarely true; 3 Sometimes true; 4 Often true; 5 Very often true.",
        "reverse_scoring_markers": "Supports reverse scoring markers: (R) or (反向)",
        "multiple_scoring_ranges": "Supports multiple scoring ranges: 1-5, 1-7, 1-6, etc. (auto-detected)",
        "subject_background_file_title": "Subject Background File:",
        "supported_formats": "Supported formats: Excel (.xlsx, .xls) and CSV (.csv)",
        "mandatory_columns": "Mandatory columns: 被试ID/Subject ID, 年龄/Age, 性别/Gender",
        "other_columns": "Other columns will be automatically parsed and added to prompts",
        "null_value_handling": "Null values will be filled with 'Not Applicable' and a missing value report will be generated",
        "high_missing_value_handling": "If any column has more than 20% missing values, a popup will notify and ask the user whether to continue",
        "usage_guide_title": "Usage Guide:",
        "usage_step_1": "After running the script, read and agree to the agreement in the welcome window",
        "usage_step_2": "Configure API parameters in the settings window (if needed)",
        "usage_step_3": "Select whether to enable random question order and consecutive dimension limit",
        "usage_step_4": "Select questionnaire file (Excel, CSV, or Word format)",
        "usage_step_5": "Select subject background file (Excel format)",
        "usage_step_6": "Select output result path",
        "usage_step_7": "Click the 'Start Processing' button to begin processing",
        "license_title": "GNU GENERAL PUBLIC LICENSE",
        "license_version": "Version 3, 29 June 2007",
        "license_paragraph_1": "This program is free software: you can redistribute it and/or modify",
        "license_paragraph_2": "it under the terms of the GNU General Public License as published by",
        "license_paragraph_3": "the Free Software Foundation, either version 3 of the License, or",
        "license_paragraph_4": "(at your option) any later version.",
        "license_paragraph_5": "You should have received a copy of the GNU General Public License",
        "license_website": "along with this program. If not, see <https://www.gnu.org/licenses/>.",
        "missing_mandatory_columns": "Background file missing mandatory columns",
        "no_valid_subjects_loaded": "No valid subjects loaded, please check and try again",
        "warning_high_missing_values": "Warning: High Missing Values",
        "high_missing_value_columns": "The following columns have more than 20% missing values:",
        "continue_running": "Continue running?",
        "warning_age_range": "Warning: Age Range Check",
        "age_out_of_range": "Found",
        "subjects": "subjects with age outside the set range",
        "years": "years",
        "invalid_ages": "Invalid ages",
        "continue_with_filter": "Continue running (subjects with invalid ages will be automatically filtered out)?",
        "scoring_rules_title": "Scoring Rules Settings",
        "score_range_recognition_rule": "Score Range Recognition Rule",
        "auto_recognize_score_range": "Automatically recognize score range: determined by numbers in scoring criteria",
        "example_score_range_1": "Example: contains '7' means 1-7 point scoring, contains '6' means 1-6 point scoring",
        "default_score_range": "Default is 1-5 point scoring",
        "reverse_scoring_rule": "Reverse Scoring Rule",
        "auto_recognize_reverse_markers": "Automatically recognize reverse scoring markers: (R) or (反向)",
        "reverse_scoring_calculation": "Reverse scoring calculation: (minimum + maximum) - original score",
        "example_reverse_scoring_1": "Example: In 5-point scoring, original score 1 becomes 5",
        "example_reverse_scoring_2": "Example: In 7-point scoring, original score 2 becomes 6",
        "dimension_score_calculation_rule": "Dimension Score Calculation Rule",
        "dimension_score_sum": "Dimension score = sum of all question scores in the dimension",
        "missing_value_handling": "Supports missing value handling: use mean replacement when only one missing value",
        "no_missing_values": "Direct sum when no missing values",
        "custom_rule_examples": "Custom Rule Examples",
        "example_modify_score_range": "Example 1: Modify score range recognition",
        "force_5_point_scoring": "Force 5-point scoring",
        "example_modify_reverse_calculation": "Example 2: Modify reverse scoring calculation",
        "example_modify_dimension_calculation": "Example 3: Modify dimension score calculation to average",
        "strategy_keep_original": "Keep Original",
        "strategy_keep_original_desc": "Keep Chinese field names as they are (recommended, AI can understand)",
        "strategy_auto_translate": "Auto Translate",
        "strategy_auto_translate_desc": "Try to automatically translate to English",
        "strategy_pinyin": "Pinyin Conversion",
        "strategy_pinyin_desc": "Use pinyin as English identifiers",
        "strategy_custom_map": "Custom Mapping",
        "strategy_custom_map_desc": "Manually specify English names when editing prompt template",
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
        "pause_button": "Pause",
        "resume_button": "Resume",
        "cancel_button": "Cancel",
        
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