#!/usr/bin/env bash
# HTTPS(443 포트) 방화벽 개방 — 사용자가 직접 실행:
#   bash ~/hr-chatbot/deploy/open_https.sh
export CLOUDSDK_PYTHON=/Users/wonholee/.local/share/uv/python/cpython-3.12-macos-aarch64-none/bin/python3.12
~/google-cloud-sdk/bin/gcloud compute firewall-rules create allow-https \
  --project project-20c05d21-327a-4f6d-b3a \
  --allow tcp:443 \
  --source-ranges 0.0.0.0/0 \
  --description "HTTPS for wonho98.duckdns.org portal"
echo "완료! https://wonho98.duckdns.org 접속 가능"
