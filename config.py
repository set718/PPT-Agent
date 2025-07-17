#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
配置管理模块
统一管理项目配置参数
"""

import os
from dataclasses import dataclass, field
from typing import Optional, Dict, Any

@dataclass
class Config:
    """项目配置类"""
    
    # API配置
    deepseek_api_key: Optional[str] = None
    deepseek_base_url: str = "https://api.deepseek.com"
    
    # PPT模板配置
    default_ppt_template: str = r"D:\jiayihan\Desktop\ppt format V1_2.pptx"
    
    # 输出配置
    output_dir: str = "output"
    temp_output_dir: str = "temp_output"
    
    # AI配置
    ai_model: str = "deepseek-chat"
    ai_temperature: float = 0.3
    ai_max_tokens: int = 2000
    
    # 文件处理配置
    max_file_size_mb: int = 50
    supported_formats: list = field(default_factory=lambda: ['.pptx'])
    
    # 日志配置
    log_level: str = "INFO"
    log_file: str = "app.log"
    
    # Web界面配置
    web_title: str = "文本转PPT填充器"
    web_icon: str = "📊"
    web_layout: str = "wide"
    
    def __post_init__(self):
        """初始化后处理"""
        # 从环境变量获取API密钥
        if not self.deepseek_api_key:
            self.deepseek_api_key = os.getenv('DEEPSEEK_API_KEY')
        
        # 创建输出目录
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.temp_output_dir, exist_ok=True)
    
    def validate(self) -> Dict[str, Any]:
        """验证配置有效性"""
        errors = {}
        warnings = {}
        
        # 检查API密钥
        if not self.deepseek_api_key:
            errors['deepseek_api_key'] = "API密钥未设置"
        
        # 检查PPT模板文件
        if not os.path.exists(self.default_ppt_template):
            warnings['default_ppt_template'] = f"默认PPT模板不存在: {self.default_ppt_template}"
        
        # 检查输出目录权限
        for directory in [self.output_dir, self.temp_output_dir]:
            if not os.path.exists(directory):
                try:
                    os.makedirs(directory, exist_ok=True)
                except Exception as e:
                    errors[f'{directory}_permission'] = f"无法创建目录 {directory}: {e}"
        
        return {
            'errors': errors,
            'warnings': warnings,
            'is_valid': len(errors) == 0
        }
    
    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return {
            'deepseek_api_key': '***' if self.deepseek_api_key else None,
            'deepseek_base_url': self.deepseek_base_url,
            'default_ppt_template': self.default_ppt_template,
            'output_dir': self.output_dir,
            'temp_output_dir': self.temp_output_dir,
            'ai_model': self.ai_model,
            'ai_temperature': self.ai_temperature,
            'ai_max_tokens': self.ai_max_tokens,
            'max_file_size_mb': self.max_file_size_mb,
            'supported_formats': self.supported_formats,
            'log_level': self.log_level,
            'log_file': self.log_file,
            'web_title': self.web_title,
            'web_icon': self.web_icon,
            'web_layout': self.web_layout
        }

# 全局配置实例
config = Config()

def get_config() -> Config:
    """获取配置实例"""
    return config

def update_config(**kwargs) -> None:
    """更新配置"""
    global config
    for key, value in kwargs.items():
        if hasattr(config, key):
            setattr(config, key, value)
        else:
            raise ValueError(f"未知的配置项: {key}")

def load_config_from_file(file_path: str) -> None:
    """从文件加载配置"""
    if not os.path.exists(file_path):
        return
    
    try:
        import json
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        global config
        for key, value in data.items():
            if hasattr(config, key):
                setattr(config, key, value)
    except Exception as e:
        print(f"加载配置文件失败: {e}")

def save_config_to_file(file_path: str) -> None:
    """保存配置到文件"""
    try:
        import json
        data = config.to_dict()
        # 不保存敏感信息
        data.pop('deepseek_api_key', None)
        
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"保存配置文件失败: {e}")

# 在导入时尝试加载配置文件
load_config_from_file('config.json')