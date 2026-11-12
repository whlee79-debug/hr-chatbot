#!/usr/bin/env bash
# HR 챗봇 외부 접속용 방화벽 개방 (8000 포트)
# 사용자가 직접 실행: bash ~/hr-chatbot/deploy/open_firewall.sh
export CLOUDSDK_PYTHON=/Users/wonholee/.local/share/uv/python/cpython-3.12-macos-aarch64-none/bin/python3.12
~/google-cloud-sdk/bin/gcloud compute firewall-rules create allow-hr-chatbot \
  --project project-20c05d21-327a-4f6d-b3a \
  --allow tcp:8000 \
  --source-ranges 0.0.0.0/0 \
  --description "HR chatbot web access (password-protected app)"
echo "완료! http://34.83.67.147:8000 접속 가능"
