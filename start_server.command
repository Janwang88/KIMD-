#!/bin/bash
cd "$(dirname "$0")"
echo "正在启动物料查询工具服务器..."
echo "=================================="
echo "请勿关闭此窗口"
echo "如果不小心关闭，请双击重新运行"
echo "=================================="

# 检查 node 是否安装
if ! command -v node &> /dev/null; then
    echo "❌ 错误: 未找到 Node.js"
    echo "请先安装 Node.js: https://nodejs.org/"
    read -p "按回车键退出..."
    exit 1
fi

# 安装依赖（如果需要）
if [ ! -d "node_modules" ]; then
    echo "📦 正在安装依赖..."
    npm install
fi

# 启动服务器
node server.js
