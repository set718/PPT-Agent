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
    openai_base_url: str = "https://openrouter.ai/api/v1"
    openai_api_key: str = ""
    
    # PPT模板配置
    default_ppt_template: str = os.path.join(os.path.dirname(__file__), "templates", "ppt_template.pptx")
    
    # 输出配置 - 云端部署时使用相对路径
    output_dir: str = "output"
    temp_output_dir: str = "temp_output"
    
    # AI配置
    ai_model: str = "gpt-4o"
    ai_temperature: float = 0.3
    ai_max_tokens: int = 2000
    
    # 模型选择配置
    available_models: Dict[str, Dict[str, Any]] = field(default_factory=lambda: {
        "gpt-4o": {
            "name": "GPT-4o",
            "description": "OpenAI GPT-4o模型，支持视觉分析功能",
            "supports_vision": True,
            "cost": "较高",
            "base_url": "https://openrouter.ai/api/v1",
            "api_provider": "OpenRouter",
            "api_key_url": "https://openrouter.ai/keys"
        },
        "deepseek-chat": {
            "name": "DeepSeek Chat",
            "description": "DeepSeek Chat模型，成本较低但不支持视觉分析",
            "supports_vision": False,
            "cost": "较低",
            "base_url": "https://api.deepseek.com/v1",
            "api_provider": "DeepSeek",
            "api_key_url": "https://platform.deepseek.com/api_keys"
        }
    })
    
    # 根据模型自动启用/禁用视觉分析
    enable_visual_analysis: bool = True
    
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
    
    # PPT布局配置
    layout_margins: Dict[str, float] = field(default_factory=lambda: {
        'slide_margin_left': 0.5,      # 幻灯片左边距（英寸）
        'slide_margin_right': 0.5,     # 幻灯片右边距（英寸）
        'slide_margin_top': 1.5,       # 幻灯片上边距（英寸）
        'slide_margin_bottom': 0.5,    # 幻灯片下边距（英寸）
        'shape_spacing': 0.1,          # 形状间距（英寸）
        'shape_margin': 0.1,           # 形状内边距（英寸）
    })
    
    # 字体大小配置
    font_sizes: Dict[str, int] = field(default_factory=lambda: {
        'large_area': 14,   # 大区域字体大小（磅）
        'medium_area': 12,  # 中等区域字体大小（磅）
        'small_area': 10,   # 小区域字体大小（磅）
        'default': 16,      # 默认字体大小（磅）
    })
    
    # 区域阈值配置
    layout_thresholds: Dict[str, float] = field(default_factory=lambda: {
        'large_area': 2.0,  # 大区域阈值（平方英寸）
        'medium_area': 1.0, # 中等区域阈值（平方英寸）
    })
    
    def __post_init__(self):
        """初始化后处理"""
        
        # 如果默认模板路径不存在，尝试查找其他可能的位置
        if not os.path.exists(self.default_ppt_template):
            possible_paths = [
                r"D:\jiayihan\Desktop\ppt format V1_2.pptx",  # 原始位置
                os.path.join(os.getcwd(), "ppt format V1_2.pptx"),  # 当前目录
                os.path.join(os.path.dirname(__file__), "templates", "ppt_template.pptx"),  # 相对于脚本位置
                os.path.join(os.path.dirname(__file__), "ppt format V1_2.pptx"),  # 脚本同级目录
            ]
            
            for path in possible_paths:
                if os.path.exists(path):
                    self.default_ppt_template = path
                    break
        
        # 创建输出目录
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.temp_output_dir, exist_ok=True)
        
        # 创建模板目录
        template_dir = os.path.dirname(self.default_ppt_template)
        if template_dir and not os.path.exists(template_dir):
            try:
                os.makedirs(template_dir, exist_ok=True)
            except OSError:
                pass  # 无法创建目录，稍后在验证中处理
        
        # 根据当前选择的模型自动设置视觉分析功能
        self._update_visual_analysis_setting()
    
    def _update_visual_analysis_setting(self):
        """根据当前模型自动设置视觉分析功能"""
        if self.ai_model in self.available_models:
            model_info = self.available_models[self.ai_model]
            self.enable_visual_analysis = model_info.get('supports_vision', False)
    
    def set_model(self, model_name: str):
        """设置AI模型并自动更新相关设置"""
        if model_name in self.available_models:
            self.ai_model = model_name
            self._update_visual_analysis_setting()
        else:
            raise ValueError(f"不支持的模型: {model_name}。支持的模型: {list(self.available_models.keys())}")
    
    def get_model_info(self, model_name: Optional[str] = None) -> Dict[str, Any]:
        """获取模型信息"""
        model = model_name if model_name is not None else self.ai_model
        return self.available_models.get(model, {})
    
    def validate(self) -> Dict[str, Any]:
        """验证配置有效性"""
        errors = {}
        warnings = {}
        
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
            'openai_base_url': self.openai_base_url,
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
    except json.JSONDecodeError as e:
        print(f"配置文件JSON格式错误: {e}")
    except FileNotFoundError:
        print(f"配置文件不存在: {file_path}")
    except PermissionError:
        print(f"无权限访问配置文件: {file_path}")
    except Exception as e:
        print(f"加载配置文件时发生未知错误: {e}")

def save_config_to_file(file_path: str) -> None:
    """保存配置到文件"""
    try:
        import json
        data = config.to_dict()
        # 配置文件不包含敏感信息
        
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except PermissionError:
        print(f"无权限写入配置文件: {file_path}")
    except OSError as e:
        print(f"文件系统错误: {e}")
    except Exception as e:
        print(f"保存配置文件时发生未知错误: {e}")

# 在导入时尝试加载配置文件
load_config_from_file('config.json')