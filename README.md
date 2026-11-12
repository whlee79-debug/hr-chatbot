# HR 규정 상담 챗봇 (hr-chatbot)

취업규칙·복무규정·근로기준법 등 규정 문서(PDF/DOCX)를 근거로, 조항을 인용해 답하는 사내 웹 챗봇입니다.
경영 대시보드 포털 [biz-portal](https://github.com/whlee79-debug/biz-portal)의 `/chatbot/`으로 서비스됩니다.
**회사 규정 문서는 가상의 회사 "홍길동" 기준으로 익명화한 샘플입니다.** (법령 문서는 공개 법령 원문)

- **라이브**: https://wonho98.duckdns.org/chatbot/
  - DuckDNS 무료 도메인 + Let's Encrypt HTTPS(자동 갱신) + GCP VM(Nginx 리버스 프록시)
- 포트폴리오 데모라 접속 비밀번호는 꺼 두었습니다 (`APP_PASSWORD`를 설정하면 켜짐)
- AI 백엔드: **DeepSeek**(기본, OpenAI 호환 API) / Claude / 데모(키워드 검색) 3중 모드
- 답변은 SSE로 토큰 단위 스트리밍, 조문 검색은 RAGFlow(설정 시) → 키워드 검색 폴백

## 구조

```
hr-chatbot/
├── app.py               # FastAPI 백엔드
├── static/index.html    # 로그인 + 챗 화면 (단일 HTML)
├── policies/            # ★ 규정 문서 (PDF/DOCX)
├── deploy/
│   ├── setup_vm.sh      # VM에 systemd 서비스로 설치하는 스크립트
│   └── open_firewall.sh # GCP 방화벽 8000 포트 개방
├── render.yaml          # (대안) Render 배포 설정
└── requirements.txt
```

## 동작 모드

`.env`(또는 환경변수) 설정에 따라 자동 선택됩니다:

| 우선순위 | 조건 | 모드 |
|---|---|---|
| 1 | `AI_API_KEY` 있음 | **OpenAI 호환 API** (기본값 DeepSeek `deepseek-chat`) |
| 2 | Anthropic 인증 있음 | Claude (문서 원문 인용 기능 포함) |
| 3 | 없음 | 데모 — AI 없이 규정 조항 키워드 검색 |

`.env` 예시:
```
AI_API_KEY=sk-...        # DeepSeek API 키
APP_PASSWORD=변경하세요   # 직원 공유용 접속 비밀번호
# AI_BASE_URL / AI_MODEL 로 GLM, Qwen 등 다른 OpenAI 호환 API 사용 가능
```

## 로컬 실행

```bash
cd hr-chatbot
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
# .env 파일 작성 (위 예시 참고) + policies/ 에 규정 문서 넣기
uvicorn app:app --host 0.0.0.0 --port 8000
```

브라우저에서 http://localhost:8000 접속.

```bash
# 동작 확인 (비밀번호 없이 status/mode만, 비밀번호 헤더 추가 시 문서 목록까지)
curl http://localhost:8000/api/health
```

## 현재 배포: GCP Compute Engine VM

기존 n8n용 e2-micro 인스턴스(`n8n-instance`, us-west1-b)에 함께 배포되어 있습니다.
**추가 비용 0원** (기존 VM에 얹음), 고정 IP `34.83.67.147` (VM을 중지해도 안 바뀜).

- systemd 서비스 `hr-chatbot` 으로 등록: 죽으면 5초 후 자동 재시작, 재부팅 시 자동 시작
- 메모리 400MB 상한(`MemoryMax`)으로 n8n과 공존 (VM에 스왑 2GB 추가되어 있음)
- 코드 위치: VM의 `/opt/hr-chatbot`, 환경설정: `/opt/hr-chatbot/.env`

### 규정 문서 갱신 방법

```bash
# 로컬 policies/ 수정 후, 문서와 코드를 VM에 반영:
gcloud compute scp -r policies app.py static/index.html \
  n8n-instance:/tmp/hr-update --zone us-west1-b
gcloud compute ssh n8n-instance --zone us-west1-b --command \
  "sudo cp -r /tmp/hr-update/* /opt/hr-chatbot/ && sudo systemctl restart hr-chatbot"
```
재시작 후 문서 로딩에 2~3분 걸립니다 (e2-micro).

### 서비스 점검/로그

```bash
gcloud compute ssh n8n-instance --zone us-west1-b --command \
  "systemctl status hr-chatbot; sudo journalctl -u hr-chatbot -n 30"
```

### 비밀번호 변경

VM의 `/opt/hr-chatbot/.env`에서 `APP_PASSWORD` 수정 → `sudo systemctl restart hr-chatbot`
(직원들은 챗 화면 우상단 "비밀번호 재설정"으로 새 비밀번호 재입력)

## 대안 배포: Render 무료 티어

GCP VM을 쓰지 않을 경우의 대안입니다. `render.yaml`이 준비되어 있습니다.

1. [render.com](https://render.com) 가입 (GitHub 로그인) → New → Web Service → 이 저장소 선택
2. 환경변수 `AI_API_KEY`, `APP_PASSWORD` 입력
3. `https://hr-chatbot-xxxx.onrender.com` 주소 발급

특성: 월 750시간 무료, 카드 불필요. 단 **15분 미사용 시 잠들고 깨는 데 30~50초**
(cron-job.org 등으로 10분마다 `/api/health` 호출하면 상시 유지 가능).
git push 시 자동 재배포되는 장점이 있습니다.

## 운영 체크리스트

1. **인증**: `APP_PASSWORD` 설정 시 접속 비밀번호 적용 (데모 사이트는 꺼 둠)
2. **HTTPS**: ✅ Nginx + Let's Encrypt 적용됨 (자동 갱신).
3. **API 키 관리**: `.env`는 git에서 제외됨. 키가 노출되면 DeepSeek 콘솔에서 재발급.
4. **로그**: 질문/답변 내용은 저장하지 않음 (토큰 사용량만 서버 로그에 기록).
5. **비용**: DeepSeek 기준 질문당 수십 원 수준, 캐시 적중 시 몇 원.
   서버 로그의 `usage(...)` 라인으로 토큰 사용량 확인.
6. **문서 관리 담당자 지정**: 규정 개정 시 위 "규정 문서 갱신 방법" 절차 수행.
7. **주의**: 규정 문서와 질문이 DeepSeek(중국 소재) 서버로 전송됩니다.
   민감도에 따라 Claude 모드 전환을 검토하세요.
