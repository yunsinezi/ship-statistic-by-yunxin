#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Railway 部署启动文件
用于在生产环境中启动 Flask 应用
"""

import os
from app import app

if __name__ == "__main__":
    # Railway 会设置 PORT 环境变量
    port = int(os.environ.get("PORT", 5000))
    
    # 生产环境配置
    debug = os.environ.get("FLASK_DEBUG", "False").lower() == "true"
    
    print(f"Starting Flask app on port {port}...")
    print(f"Debug mode: {debug}")
    
    # 启动应用
    app.run(
        host="0.0.0.0",  # 必须监听所有接口
        port=port,
        debug=debug,
        use_reloader=False
    )
