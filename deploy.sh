#!/bin/bash
# 裕安走访系统 一键部署脚本（Ubuntu 22.04+，root运行）
set -e
echo "=== 裕安走访系统 云端部署开始 ==="

echo "[1/6] 安装依赖..."
export DEBIAN_FRONTEND=noninteractive
apt-get update -qq
apt-get install -y -qq python3 python3-pip curl > /dev/null

echo "[2/6] 安装Python库..."
pip3 install -q openpyxl 2>/dev/null || pip3 install -q --break-system-packages openpyxl

echo "[3/6] 创建目录..."
mkdir -p /opt/visit
cd /opt/visit

echo "[4/6] 下载服务程序..."
curl -sSL -o /opt/visit/cloud_server.py "https://jeremywuuuuu.github.io/visit/cloud_server.py"
if [ ! -s /opt/visit/cloud_server.py ]; then
  echo "下载失败，请检查网络后重试"
  exit 1
fi
python3 -c "import ast;ast.parse(open('/opt/visit/cloud_server.py',encoding='utf-8').read())" || { echo "程序校验失败"; exit 1; }

echo "[5/6] 配置开机自启..."
cat > /etc/systemd/system/visit.service <<'EOF'
[Unit]
Description=Yunan Visit System
After=network.target
[Service]
Type=simple
ExecStart=/usr/bin/python3 /opt/visit/cloud_server.py
WorkingDirectory=/opt/visit
Restart=always
RestartSec=5
[Install]
WantedBy=multi-user.target
EOF

systemctl daemon-reload
systemctl enable visit > /dev/null 2>&1
systemctl restart visit

echo "[6/6] 验证..."
sleep 3
IP=$(curl -s ifconfig.me || echo "服务器IP")
if curl -s -o /dev/null -w "%{http_code}" "http://127.0.0.1/api/results" | grep -q 200; then
  echo ""
  echo "========================================"
  echo "  部署成功！"
  echo "  访问地址: http://$IP/"
  echo "  管理总览: http://$IP/visit.html?k=admf01jp58156"
  echo "========================================"
else
  echo "服务启动异常，请执行: journalctl -u visit -n 30"
  exit 1
fi
