import os
import sys

# 将根目录添加到 sys.path 以便测试能顺利引入 app_ui, pipeline 等模块
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
