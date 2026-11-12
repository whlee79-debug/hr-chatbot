#!/usr/bin/env bash
# GCP VM에서 HR 챗봇을 systemd 서비스로 설치하는 스크립트.
# 사용법: 프로젝트를 /opt/hr-chatbot 에 풀어놓은 뒤  sudo bash deploy/setup_vm.sh
set -euo pipefail

APP_DIR=/opt/hr-chatbot

echo "[1/4] 파이썬 가상환경 준비"
apt-get update -qq
apt-get install -y -qq python3-venv python3-pip > /dev/null
cd "$APP_DIR"
python3 -m venv .venv
.venv/bin/pip install -q --upgrade pip
.venv/bin/pip install -q -r requirements.txt

echo "[2/4] .env 확인"
if [ ! -f "$APP_DIR/.env" ]; then
  echo "경고: $APP_DIR/.env 가 없습니다. AI_API_KEY / APP_PASSWORD 를 넣어주세요."
fi
chmod 600 "$APP_DIR/.env" 2>/dev/null || true

echo "[3/4] systemd 서비스 등록"
cat > /etc/systemd/system/hr-chatbot.service <<'UNIT'
[Unit]
Description=HR Policy Chatbot (FastAPI)
After=network.target

[Service]
WorkingDirectory=/opt/hr-chatbot
ExecStart=/opt/hr-chatbot/.venv/bin/uvicorn app:app --host 0.0.0.0 --port 8000
Restart=always
RestartSec=5
# 메모리 상한 (e2-micro 보호: n8n과 공존을 위해 400MB로 제한)
MemoryMax=400M

[Install]
WantedBy=multi-user.target
UNIT

systemctl daemon-reload
systemctl enable hr-chatbot
systemctl restart hr-chatbot

echo "[4/4] 상태 확인"
sleep 3
systemctl --no-pager status hr-chatbot | head -8
echo
echo "완료! 다음으로 확인하세요:  curl -s http://localhost:8000/api/health"
