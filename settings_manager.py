"""
用户设置管理器
用于保存和加载用户的设置和选择
"""

import json
import os
from pathlib import Path

# 设置文件路径
SETTINGS_FILE = Path(__file__).parent / "user_settings.json"

def load_settings():
    """加载用户设置"""
    try:
        if SETTINGS_FILE.exists():
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        else:
            # 返回默认设置
            return {
                "language": "zh",
                "welcome_shown": False,
                "api_settings": {
                    "api_key": "",
                    "base_url": "",
                    "model_name": ""
                },
                "questionnaire_settings": {
                    "random_order": False,
                    "max_consecutive_same_dim": 3,
                    "token_limit": 4000,
                    "max_tokens_per_response": 512,
                    "min_age": 18,
                    "max_age": 75
                },
                "file_selection": {
                    "questionnaire_file": "",
                    "background_file": ""
                },
                "output_settings": {
                    "output_format": "excel",
                    "output_filename": "EasyPsych_Results",
                    "output_directory": ""
                }
            }
    except Exception as e:
        print(f"加载设置文件时出错: {e}")
        # 返回默认设置
        return {
            "language": "zh",
            "welcome_shown": False,
            "api_settings": {},
            "questionnaire_settings": {},
            "file_selection": {},
            "output_settings": {}
        }

def save_settings(settings):
    """保存用户设置"""
    try:
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False, indent=4)
        return True
    except Exception as e:
        print(f"保存设置文件时出错: {e}")
        return False

def update_setting(key, value):
    """更新单个设置项"""
    settings = load_settings()
    
    # 支持嵌套键，如 "api_settings.api_key"
    keys = key.split('.')
    current = settings
    
    for k in keys[:-1]:
        if k not in current:
            current[k] = {}
        current = current[k]
    
    current[keys[-1]] = value
    return save_settings(settings)

def get_setting(key, default=None):
    """获取单个设置项"""
    settings = load_settings()
    
    # 支持嵌套键，如 "api_settings.api_key"
    keys = key.split('.')
    current = settings
    
    for k in keys:
        if isinstance(current, dict) and k in current:
            current = current[k]
        else:
            return default
    
    return current

# 语言设置相关的便捷函数
def get_language():
    """获取用户选择的语言"""
    return get_setting("language", "zh")

def set_language(lang):
    """设置用户选择的语言"""
    return update_setting("language", lang)

def is_welcome_shown():
    """检查是否已经显示过欢迎界面"""
    return get_setting("welcome_shown", False)

def set_welcome_shown(shown=True):
    """设置欢迎界面显示状态"""
    return update_setting("welcome_shown", shown)