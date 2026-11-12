"""사내 HR 정책 상담 챗봇 — FastAPI 백엔드.

policies/ 폴더의 PDF 규정 문서를 Claude Files API에 업로드해 두고,
직원 질문마다 문서 전체를 컨텍스트로 제공한다 (인용 + 프롬프트 캐싱).
"""

import hashlib
import hmac
import json
import logging
import os
import re
from pathlib import Path

import anthropic
import docx
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from openai import OpenAI
from pypdf import PdfReader
from fastapi import FastAPI, Header, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("hr-chatbot")

BASE_DIR = Path(__file__).resolve().parent
POLICY_DIR = BASE_DIR / "policies"
FILE_ID_CACHE = BASE_DIR / "file_ids.json"

MODEL = "claude-opus-4-8"
BETAS = ["files-api-2025-04-14"]

def _load_dotenv() -> None:
    """BASE_DIR/.env 의 KEY=VALUE 를 환경변수로 로드 (이미 설정된 값은 유지)."""
    env_file = BASE_DIR / ".env"
    if not env_file.exists():
        return
    for line in env_file.read_text().splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            key, _, value = line.partition("=")
            os.environ.setdefault(key.strip(), value.strip().strip("\"'"))


_load_dotenv()

# OpenAI 호환 API 설정 (.env 또는 환경변수)
# AI_API_KEY 가 설정되어 있으면 Claude 대신 이 API를 사용한다.
# 기본값은 DeepSeek. GLM/Qwen 등도 BASE_URL/MODEL만 바꾸면 동작:
#   DeepSeek : https://api.deepseek.com                        / deepseek-chat
#   Zhipu GLM: https://open.bigmodel.cn/api/paas/v4            / glm-4-flash
#   Qwen     : https://dashscope-intl.aliyuncs.com/compatible-mode/v1 / qwen-flash
AI_API_KEY = os.environ.get("AI_API_KEY", "")
AI_BASE_URL = os.environ.get("AI_BASE_URL", "https://api.deepseek.com")
AI_MODEL = os.environ.get("AI_MODEL", "deepseek-chat")
# NIM(NVIDIA 등)에 호스팅된 deepseek 계열 추론 모델의 thinking 모드를 끄는 옵션.
# 기본은 꺼짐(빈 값) — api.deepseek.com 등은 이 파라미터를 모르므로 안전하게 opt-in.
# NIM 배포에서 속도를 원하면 .env 에 AI_DISABLE_THINKING=1 설정.
AI_DISABLE_THINKING = os.environ.get("AI_DISABLE_THINKING", "0") not in ("0", "false", "False", "")
AI_EXTRA_BODY = {"chat_template_kwargs": {"thinking": False}} if AI_DISABLE_THINKING else {}

# 접속 비밀번호 (인터넷 공개 배포 시 필수). 비어 있으면 인증 없이 동작(로컬 개발용).
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")

SYSTEM_PROMPT = """당신은 사내 HR 정책 상담 도우미입니다.

규칙:
- 반드시 제공된 규정 문서(취업규칙, 복무규정, 휴가규정 등)에 근거해서만 답변합니다.
- 답변할 때 근거가 되는 규정명과 조항(제○조)을 함께 언급합니다.
- 문서에 없는 사항은 추측하지 말고 "해당 내용은 규정 문서에서 확인되지 않습니다. HR 담당자에게 문의해 주세요."라고 안내합니다.
- 징계, 급여 이의, 직장 내 괴롭힘 신고, 개인별 예외 적용 등 민감하거나 개별 판단이 필요한 사안은 직접 답하지 말고 HR 담당자에게 문의하도록 안내합니다.
- 답변은 한국어로, 간결하고 친절하게 작성합니다.
- 본 답변은 참고용이며 법적 효력이 없다는 점을 필요한 경우 알립니다.

규정과 법령이 충돌할 때의 해석 기준:
1. 원칙적으로 회사 규정(취업규칙 등)을 우선 적용합니다.
2. 다만 회사 규정이 법령(근로기준법 등)보다 근로자에게 불리한 경우에는 법령을 우선 적용합니다.
   (법령은 최저 기준이므로, 회사 규정은 법령 이상으로만 유효합니다.)
- 회사 규정과 법령의 내용이 다를 때는 어느 쪽이 근로자에게 유리한지 판단해 위 기준을 적용하고,
  판단 근거(규정/법령 조항)를 함께 밝힙니다. 판단이 애매하면 HR 담당자에게 확인하도록 안내합니다.

사내 규정끼리 내용이 다를 때의 해석 기준:
1. 취업규칙이 상위 규정이므로 휴가 일수·요건 등 근로조건은 취업규칙을 우선 적용합니다.
2. 세부 규정(경조휴가 규정 등)에만 있는 항목(경조비, 선물 등 금전적 지원)은 취업규칙과 충돌이 아니라
   보완이므로 그대로 적용해 안내합니다.
3. 같은 항목의 수치가 서로 다르면(예: 휴가 일수) 두 규정의 내용을 각각 출처와 함께 제시하고,
   위 1번 기준에 따른 결론을 명확히 한 뒤, 규정 정비가 필요할 수 있음을 덧붙입니다.
4. 서로 모순되는 설명을 한 답변 안에 섞지 않습니다. 결론을 먼저 한 문장으로 말하고 근거를 정리합니다."""

client = anthropic.Anthropic()  # ANTHROPIC_API_KEY 환경변수 사용


def _has_credentials() -> bool:
    """API 인증 수단(환경변수 키 또는 ant 로그인 프로필)이 있는지."""
    if os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN"):
        return True
    profile_dir = Path.home() / ".config" / "anthropic" / "credentials"
    return profile_dir.exists() and any(profile_dir.iterdir())


def _resolve_mode() -> str:
    """동작 모드: openai(OpenAI 호환 API) > claude > demo(키워드 검색)."""
    if AI_API_KEY:
        return "openai"
    if _has_credentials():
        return "claude"
    return "demo"


MODE = _resolve_mode()

openai_client = OpenAI(api_key=AI_API_KEY or "unused", base_url=AI_BASE_URL)

app = FastAPI(title="HR 정책 상담 챗봇")

# 규정 문서 목록 — 서버 시작 시 채워진다
# {"title": str, "kind": "pdf"|"text", "file_id": str} 또는 {"title", "kind", "text"}
policy_documents: list[dict] = []


def _extract_docx_text(path: Path) -> str:
    """docx에서 문단과 표를 문서 순서대로 텍스트로 추출한다."""
    document = docx.Document(str(path))
    parts: list[str] = []
    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            text = Paragraph(child, document).text.strip()
            if text:
                parts.append(text)
        elif child.tag == qn("w:tbl"):
            for row in Table(child, document).rows:
                cells = [c.text.strip().replace("\n", " ") for c in row.cells]
                if any(cells):
                    parts.append(" | ".join(cells))
    return "\n".join(parts)


def _extract_pdf_text(path: Path) -> str:
    """PDF에서 텍스트를 추출한다. 스캔본(이미지) PDF는 빈 문자열이 나올 수 있다."""
    try:
        reader = PdfReader(str(path))
        return "\n".join((page.extract_text() or "") for page in reader.pages).strip()
    except Exception as e:
        log.warning("PDF 텍스트 추출 실패(%s): %s", path.name, e)
        return ""


def _load_file_id_cache() -> dict:
    if FILE_ID_CACHE.exists():
        return json.loads(FILE_ID_CACHE.read_text())
    return {}


def _save_file_id_cache(cache: dict) -> None:
    FILE_ID_CACHE.write_text(json.dumps(cache, ensure_ascii=False, indent=2))


@app.on_event("startup")
def load_policy_documents() -> None:
    """policies/ 의 규정 문서를 준비한다.

    - PDF: Files API에 업로드 (내용 해시로 중복 업로드 방지)
    - DOCX: 텍스트로 추출해 인라인 문서 블록으로 사용
    """
    cache = _load_file_id_cache()
    paths = sorted(
        p for p in POLICY_DIR.iterdir()
        if p.suffix.lower() in (".pdf", ".docx") and not p.name.startswith(("~$", "."))
    )
    if not paths:
        log.warning("policies/ 폴더에 PDF/DOCX가 없습니다. 규정 문서를 넣고 서버를 재시작하세요.")
        return

    for path in paths:
        title = path.stem.rstrip(".")
        if path.suffix.lower() == ".docx":
            text = _extract_docx_text(path)
            if not text.strip():
                log.warning("텍스트 추출 실패(빈 문서): %s", path.name)
                continue
            policy_documents.append({"title": title, "kind": "text", "text": text})
            log.info("텍스트 추출 완료: %s (%d자)", path.name, len(text))
            continue

        # PDF: Claude 모드는 Files API 업로드(원본 그대로, 페이지 인용 가능),
        # 그 외 모드는 텍스트를 추출해 인라인 문서로 사용
        if MODE != "claude":
            text = _extract_pdf_text(path)
            if not text:
                log.warning("텍스트 추출 실패(스캔본 PDF일 수 있음): %s — 이 문서는 제외됩니다", path.name)
                continue
            policy_documents.append({"title": title, "kind": "text", "text": text})
            log.info("PDF 텍스트 추출 완료: %s (%d자)", path.name, len(text))
            continue

        data = path.read_bytes()
        digest = hashlib.sha256(data).hexdigest()
        entry = cache.get(path.name)
        if entry and entry.get("sha256") == digest:
            file_id = entry["file_id"]
            log.info("업로드 생략(변경 없음): %s -> %s", path.name, file_id)
        else:
            try:
                uploaded = client.beta.files.upload(
                    file=(path.name, data, "application/pdf"),
                )
            except Exception as e:
                log.error("업로드 실패(%s): %s — 이 문서는 제외됩니다", path.name, e)
                continue
            file_id = uploaded.id
            cache[path.name] = {"sha256": digest, "file_id": file_id}
            log.info("업로드 완료: %s -> %s", path.name, file_id)
        policy_documents.append({"title": title, "kind": "pdf", "file_id": file_id})

    _save_file_id_cache(cache)
    log.info("규정 문서 %d건 준비 완료", len(policy_documents))


class ChatTurn(BaseModel):
    role: str  # "user" | "assistant"
    content: str


class ChatRequest(BaseModel):
    messages: list[ChatTurn]


def _build_messages(history: list[ChatTurn]) -> list[dict]:
    """대화 이력을 API 메시지로 변환. 규정 문서는 첫 user 턴에 넣는다.

    문서 블록 + 시스템 프롬프트가 매 요청 동일한 프리픽스가 되어
    프롬프트 캐싱이 적용된다 (마지막 문서 블록에 cache_control).
    """
    doc_blocks = []
    for i, doc in enumerate(policy_documents):
        if doc["kind"] == "pdf":
            source = {"type": "file", "file_id": doc["file_id"]}
        else:
            source = {"type": "text", "media_type": "text/plain", "data": doc["text"]}
        block = {
            "type": "document",
            "source": source,
            "title": doc["title"],
            "citations": {"enabled": True},
        }
        if i == len(policy_documents) - 1:
            block["cache_control"] = {"type": "ephemeral"}
        doc_blocks.append(block)

    messages: list[dict] = []
    first_user_done = False
    for turn in history:
        if turn.role == "user" and not first_user_done:
            messages.append({
                "role": "user",
                "content": doc_blocks + [{"type": "text", "text": turn.content}],
            })
            first_user_done = True
        else:
            messages.append({"role": turn.role, "content": turn.content})
    return messages


def _split_articles(text: str) -> list[str]:
    """규정 텍스트를 조항(제○조) 단위 블록으로 나눈다. 조항이 없으면 문단 묶음."""
    blocks: list[str] = []
    current: list[str] = []
    for line in text.split("\n"):
        line = line.strip()
        if not line:
            continue
        is_article_start = re.match(r"^제\s*\d+\s*조", line)
        if (is_article_start and current) or sum(len(l) for l in current) > 500:
            blocks.append("\n".join(current))
            current = []
        current.append(line)
    if current:
        blocks.append("\n".join(current))
    return blocks


def _search_excerpts(query: str, limit: int = 3) -> list[dict]:
    """질의 키워드로 규정 문서에서 관련 조항 블록을 찾아 citations 형식으로 반환."""
    raw_tokens = [t for t in re.split(r"[\s\W]+", query) if len(t) >= 2]
    # 조사가 붙은 어절도 잡히도록 2글자 접두어를 함께 사용 ("연차는" -> "연차")
    tokens = {t for t in raw_tokens} | {t[:2] for t in raw_tokens}

    scored: list[tuple[float, str, str]] = []
    seen_blocks: set[str] = set()
    for doc in policy_documents:
        if doc["kind"] != "text":
            continue
        for block in _split_articles(doc["text"]):
            if len(block) < 20 or block in seen_blocks:
                continue  # 목차 항목 등 짧은 블록과 중복 제외
            seen_blocks.add(block)
            hits = sum(block.count(t) for t in tokens)
            if hits:
                # 키워드 등장 횟수 우선, 내용이 충실한(긴) 블록에 가산점
                scored.append((hits + min(len(block), 500) / 1000, doc["title"], block))
    scored.sort(key=lambda x: -x[0])

    return [
        {"id": i + 1, "document": title, "pages": "", "cited_text": block[:400]}
        for i, (_, title, block) in enumerate(scored[:limit])
    ]


# 프롬프트 토큰을 줄여 첫 응답 지연을 낮춘다.
# - 소형 사규(이 글자 수 이하): 전문 유지 (IT자산·시차출퇴근·도서·경조 등)
# - 취업규칙 등 1순위 문서: 관련 조항을 넉넉히 발췌 + 총칙(앞 조항) 항상 포함
# - 대용량 법령·매뉴얼: 관련 조항 소수만 발췌
CONTEXT_FULL_MAX_CHARS = 6000
PRIMARY_DOC_HINTS = ("취업규칙", "Rules Of Employment")
PRIMARY_BLOCKS = 15   # 취업규칙에서 뽑을 관련 조항 수(누락 위험 최소화 위해 넉넉히)
PRIMARY_HEAD = 1      # 앞부분(목적·적용범위 등) 항상 포함할 블록 수
LEGAL_BLOCKS = 5      # 대용량 법령·매뉴얼에서 뽑을 관련 조항 수


def _query_tokens(query: str) -> tuple[set[str], set[str]]:
    """질의어를 (원형 토큰, 어간 후보) 로 나눈다.

    한국어 조사 때문에 어절이 그대로 안 맞는 경우가 많아 2~4글자 접두어도 후보로 쓰되,
    원형 토큰보다 낮은 가중치를 준다(짧은 접두어가 과매칭되는 것을 막기 위함).
    """
    raw = [t for t in re.split(r"[\s\W]+", query) if len(t) >= 2]
    # 조사(1~2글자)를 뗀 형태는 사실상 같은 단어이므로 원형과 같은 가중치로 취급
    full = set(raw) | {t[:-n] for t in raw for n in (1, 2) if len(t) - n >= 2}
    stems = {t[:n] for t in raw for n in (2, 3, 4) if len(t) > n}
    return full, stems - full


def _relevant_blocks(text: str, tokens: tuple[set[str], set[str]], limit: int,
                     always_head: int = 0) -> list[str]:
    """문서를 조항 블록으로 나눠 질문과 관련된 블록을 점수순으로 뽑는다.

    조문 제목(제○조(···)) 에 질의어가 들어가면 크게 가산한다. 산업안전보건법처럼
    문서 전체가 특정 주제어로 도배된 경우, 본문 등장 횟수만으로는 정작 핵심 조문이
    밀리기 때문이다.
    always_head > 0 이면 문서 앞부분 블록(총칙·목적 등)을 키워드와 무관하게 포함한다.
    """
    full, stems = tokens
    blocks = _split_articles(text)
    scored: list[tuple[float, str]] = []
    for block in blocks:
        if len(block) < 20:
            continue
        head = block.split("\n", 1)[0][:60]  # 조문 제목 줄
        score = 0.0
        for t in full:
            score += block.count(t) * 3.0
            if t in head:
                score += 12.0
        for t in stems:
            score += block.count(t) * 0.5
            if t in head:
                score += 3.0
        if score > 0:
            scored.append((score + min(len(block), 500) / 1000, block))
    scored.sort(key=lambda x: -x[0])
    chosen = [b for _, b in scored[:limit]]
    head = [b for b in blocks[:always_head] if len(b) >= 20]
    seen: set[str] = set()
    out: list[str] = []
    for b in head + chosen:   # 앞 조항 먼저, 이어서 관련 조항
        if b not in seen:
            seen.add(b)
            out.append(b)
    return out


# ── RAGFlow 의미 검색 (RAG_BACKEND=ragflow 일 때 사용, 실패 시 키워드 폴백) ──
RAG_BACKEND = os.getenv("RAG_BACKEND", "keyword")
RAGFLOW_URL = os.getenv("RAGFLOW_URL", "").rstrip("/")
RAGFLOW_API_KEY = os.getenv("RAGFLOW_API_KEY", "")
RAGFLOW_DATASET_IDS = [x for x in os.getenv("RAGFLOW_DATASET_IDS", "").split(",") if x]
RAGFLOW_TOP_N = int(os.getenv("RAGFLOW_TOP_N", "10"))      # 컨텍스트에 넣을 조문 수
RAGFLOW_TIMEOUT = float(os.getenv("RAGFLOW_TIMEOUT", "8")) # 초과 시 키워드 폴백


def _ragflow_docs_text(question: str) -> str | None:
    """RAGFlow 의미 검색으로 규정 컨텍스트를 만든다. 결과가 없으면 None(폴백 신호)."""
    import httpx

    resp = httpx.post(
        RAGFLOW_URL + "/api/v1/retrieval",
        headers={"Authorization": "Bearer " + RAGFLOW_API_KEY},
        json={
            "question": question,
            "dataset_ids": RAGFLOW_DATASET_IDS,
            "top_k": 128,
            "similarity_threshold": 0.15,
            "page_size": RAGFLOW_TOP_N,
        },
        timeout=RAGFLOW_TIMEOUT,
    )
    data = resp.json()
    if resp.status_code != 200 or data.get("code") != 0:
        raise RuntimeError(f"retrieval 응답 이상: {str(data)[:120]}")
    chunks = data["data"].get("chunks") or []
    if not chunks:
        return None

    # 문서별로 묶어 기존 컨텍스트와 같은 골격으로 구성 (출처 표기가 답변 인용에 쓰인다)
    by_doc: dict[str, list[str]] = {}
    total = 0
    for c in chunks:
        body = re.sub(r"\s+", " ", c.get("content", "")).strip()
        if not body:
            continue
        if total + len(body) > CONTEXT_FULL_MAX_CHARS:
            break
        by_doc.setdefault(c.get("document_keyword", "규정"), []).append(body)
        total += len(body)
    if not by_doc:
        return None
    parts = [
        f"===== 규정 문서(의미 검색 발췌): {doc} =====\n" + "\n…\n".join(bodies)
        for doc, bodies in by_doc.items()
    ]
    return (
        "----- 규정·법령 관련 조항 (의미 기반 검색으로 추림) -----\n\n"
        + "\n\n".join(parts)
    )


def _build_openai_docs_text(question: str) -> str:
    """DeepSeek 등에 넣을 규정 컨텍스트.

    RAG_BACKEND=ragflow 이면 RAGFlow 의미 검색 발췌에 키워드 발췌(축소판)를
    더한 하이브리드로 구성한다 — 이 인스턴스의 RAGFlow는 키워드 매칭이 후보를
    만들기 때문에, 일상어 질문(예: "아내가 출산")은 키워드 발췌가 안전망이 된다.
    RAGFlow 실패·무결과 시 기존 키워드 발췌 전체로 폴백한다.
    """
    if RAG_BACKEND == "ragflow" and RAGFLOW_API_KEY and RAGFLOW_DATASET_IDS:
        try:
            rag_text = _ragflow_docs_text(question)
        except Exception as exc:
            rag_text = None
            log.warning("RAGFlow 검색 실패(%s) — 키워드 발췌로 폴백", exc)
        if rag_text:
            kw_text = _keyword_docs_text(question, primary_blocks=8, legal_blocks=3)
            log.info("컨텍스트: RAGFlow %d자 + 키워드 %d자", len(rag_text), len(kw_text))
            return rag_text + "\n\n" + kw_text
        log.info("컨텍스트: 키워드 발췌 (RAGFlow 무결과/실패)")

    return _keyword_docs_text(question, PRIMARY_BLOCKS, LEGAL_BLOCKS)


def _keyword_docs_text(question: str, primary_blocks: int, legal_blocks: int) -> str:
    """키워드 스코어링 발췌 — 소형 사규는 전문, 취업규칙·법령은 관련 조항만."""
    tokens = _query_tokens(question)
    full_parts: list[str] = []
    excerpt_parts: list[str] = []
    for d in policy_documents:
        if d["kind"] != "text":
            continue
        title, text = d["title"], d["text"]
        if len(text) <= CONTEXT_FULL_MAX_CHARS:
            full_parts.append(f"===== 규정 문서: {title} =====\n{text}")
            continue
        is_primary = any(h in title for h in PRIMARY_DOC_HINTS)
        blocks = _relevant_blocks(
            text, tokens,
            primary_blocks if is_primary else legal_blocks,
            always_head=PRIMARY_HEAD if is_primary else 0,
        )
        if blocks:
            tag = "관련 조항 발췌" if is_primary else "질문 관련 발췌"
            joined = "\n…\n".join(blocks)
            excerpt_parts.append(f"===== 규정 문서({tag}): {title} =====\n{joined}")
    docs_text = "\n\n".join(full_parts)
    if excerpt_parts:
        docs_text += (
            "\n\n----- 규정·법령 관련 조항 발췌(질문과 관련된 부분만 추림) -----\n\n"
            + "\n\n".join(excerpt_parts)
        )
    return docs_text


def _demo_answer(question: str) -> dict:
    """데모 모드: AI 없이 규정 문서에서 키워드가 겹치는 조항을 찾아 보여준다."""
    citations = _search_excerpts(question)
    if citations:
        answer = (
            "🔍 [데모 모드] AI가 연결되지 않아 규정 문서에서 관련 조항을 검색해 보여드립니다.\n"
            "아래 근거 발췌를 확인해 주세요. AI 연결 후에는 이 내용을 종합한 상담 답변이 제공됩니다."
        )
    else:
        answer = (
            "🔍 [데모 모드] 규정 문서에서 관련 조항을 찾지 못했습니다.\n"
            "다른 키워드로 질문해 보시거나, HR 담당자에게 문의해 주세요."
        )
    return {
        "segments": [{"text": answer, "refs": [c["id"] for c in citations]}],
        "citations": citations,
        "citations_label": "🔎 관련 규정 발췌 (키워드 검색)",
        "demo": True,
    }


def _openai_chat(req: ChatRequest) -> dict:
    """OpenAI 호환 API(DeepSeek/GLM/Qwen 등)로 답변 생성.

    문서 블록/citations 기능이 없으므로 규정 전문을 시스템 프롬프트에 인라인하고,
    근거 조항은 답변 본문에 표기하도록 지시한다. 시스템 프롬프트가 매 요청 동일하므로
    DeepSeek 등의 자동 컨텍스트 캐싱이 적용된다.
    """
    docs_text = _build_openai_docs_text(req.messages[-1].content)
    system = (
        SYSTEM_PROMPT
        + "\n- 근거 조항은 답변 본문에 「규정명 제○조」 형식으로 표기합니다."
        + "\n- 마크다운 서식(**, ##, - 등)을 사용하지 말고 일반 텍스트로만 답변합니다. 목록이 필요하면 1. 2. 3. 번호를 사용합니다."
        + "\n\n아래는 사내 규정 문서(회사 규정은 전문, 법령·매뉴얼은 질문 관련 발췌)입니다. 반드시 이 내용에만 근거해 답변하세요.\n\n"
        + docs_text
    )
    messages = [{"role": "system", "content": system}] + [
        {"role": t.role, "content": t.content} for t in req.messages
    ]
    try:
        resp = openai_client.chat.completions.create(
            model=AI_MODEL,
            messages=messages,
            max_tokens=1200,
            extra_body=AI_EXTRA_BODY,
        )
    except Exception as e:
        log.error("OpenAI 호환 API 오류(%s): %s", AI_BASE_URL, e)
        raise HTTPException(502, "AI 응답 생성 중 오류가 발생했습니다.")

    answer = (resp.choices[0].message.content or "").strip()
    if resp.usage:
        log.info("usage(%s): prompt=%s completion=%s", AI_MODEL,
                 resp.usage.prompt_tokens, resp.usage.completion_tokens)

    # 원문 발췌 병기: 질문 + 답변 속 「인용 표기」로 관련 조항을 검색해
    # 직원이 AI 답변을 원문으로 직접 검증할 수 있게 한다.
    question = req.messages[-1].content
    cited_refs = " ".join(re.findall(r"「([^」]{2,40})」", answer))
    excerpts = _search_excerpts(f"{question} {cited_refs}")

    return {
        "segments": [{"text": answer, "refs": []}],
        "citations": excerpts,
        "citations_label": "📄 관련 규정 원문 (자동 검색 — 답변 검증용)",
    }


def _check_password(provided: str) -> None:
    if APP_PASSWORD and not hmac.compare_digest(provided or "", APP_PASSWORD):
        raise HTTPException(401, "접속 비밀번호가 올바르지 않습니다.")


class AuthRequest(BaseModel):
    password: str = ""


@app.post("/api/auth")
def auth(req: AuthRequest):
    """로그인 화면에서 비밀번호 확인용. 틀리면 401."""
    _check_password(req.password)
    return {"ok": True}


@app.post("/api/chat")
def chat(req: ChatRequest, x_app_password: str = Header(default="")):
    _check_password(x_app_password)
    if not policy_documents:
        raise HTTPException(503, "규정 문서가 아직 준비되지 않았습니다. policies/ 폴더에 PDF 또는 DOCX를 넣고 서버를 재시작하세요.")
    if not req.messages or req.messages[-1].role != "user":
        raise HTTPException(400, "마지막 메시지는 user 여야 합니다.")

    if MODE == "openai":
        return _openai_chat(req)
    if MODE == "demo":
        return _demo_answer(req.messages[-1].content)

    try:
        response = client.beta.messages.create(
            model=MODEL,
            max_tokens=2048,
            betas=BETAS,
            system=[{
                "type": "text",
                "text": SYSTEM_PROMPT,
                "cache_control": {"type": "ephemeral"},
            }],
            messages=_build_messages(req.messages),
        )
    except anthropic.RateLimitError:
        raise HTTPException(429, "요청이 많아 잠시 후 다시 시도해 주세요.")
    except anthropic.APIStatusError as e:
        log.error("Claude API 오류: %s", e)
        raise HTTPException(502, "AI 응답 생성 중 오류가 발생했습니다.")
    except anthropic.APIConnectionError:
        raise HTTPException(502, "AI 서버에 연결할 수 없습니다.")

    if response.stop_reason == "refusal":
        return {
            "segments": [{"text": "죄송합니다. 해당 질문에는 답변드릴 수 없습니다. HR 담당자에게 직접 문의해 주세요.", "refs": []}],
            "citations": [],
        }

    # 텍스트를 인용 단위(segment)로 쪼개고, 인용 목록을 별도로 반환
    segments: list[dict] = []
    citations: list[dict] = []
    seen: dict[tuple, int] = {}

    for block in response.content:
        if block.type != "text":
            continue
        refs: list[int] = []
        for c in (getattr(block, "citations", None) or []):
            title = getattr(c, "document_title", None) or "규정 문서"
            cited = (getattr(c, "cited_text", "") or "").strip()
            start = getattr(c, "start_page_number", None)
            end = getattr(c, "end_page_number", None)
            if start and end and end - start > 1:
                pages = f"{start}~{end - 1}쪽"
            elif start:
                pages = f"{start}쪽"
            else:
                pages = ""
            key = (title, cited[:120])
            if key in seen:
                idx = seen[key]
            else:
                idx = len(citations) + 1
                seen[key] = idx
                citations.append({
                    "id": idx,
                    "document": title,
                    "pages": pages,
                    "cited_text": cited[:300],
                })
            if idx not in refs:
                refs.append(idx)
        segments.append({"text": block.text, "refs": refs})

    log.info(
        "usage: input=%s cache_read=%s cache_write=%s output=%s",
        response.usage.input_tokens,
        response.usage.cache_read_input_tokens,
        response.usage.cache_creation_input_tokens,
        response.usage.output_tokens,
    )
    return {
        "segments": segments,
        "citations": citations,
        "citations_label": "📌 근거 인용 (AI가 문서 원문에서 직접 인용)",
    }


def _sse(obj: dict) -> str:
    return f"data: {json.dumps(obj, ensure_ascii=False)}\n\n"


def _openai_chat_stream(req: ChatRequest):
    """OpenAI 호환 API를 토큰 스트리밍으로 받아 SSE로 흘려보낸다.

    답변을 생성 즉시 조금씩 전달해 체감 응답 속도를 크게 줄인다.
    답변 완료 후 인용 발췌를 마지막 이벤트로 전송한다.
    """
    docs_text = _build_openai_docs_text(req.messages[-1].content)
    system = (
        SYSTEM_PROMPT
        + "\n- 근거 조항은 답변 본문에 「규정명 제○조」 형식으로 표기합니다."
        + "\n- 마크다운 서식(**, ##, - 등)을 사용하지 말고 일반 텍스트로만 답변합니다. 목록이 필요하면 1. 2. 3. 번호를 사용합니다."
        + "\n\n아래는 사내 규정 문서(회사 규정은 전문, 법령·매뉴얼은 질문 관련 발췌)입니다. 반드시 이 내용에만 근거해 답변하세요.\n\n"
        + docs_text
    )
    messages = [{"role": "system", "content": system}] + [
        {"role": t.role, "content": t.content} for t in req.messages
    ]

    def gen():
        parts: list[str] = []
        try:
            stream = openai_client.chat.completions.create(
                model=AI_MODEL, messages=messages, max_tokens=1200, stream=True,
                extra_body=AI_EXTRA_BODY,
            )
            for chunk in stream:
                delta = (chunk.choices[0].delta.content or "") if chunk.choices else ""
                if delta:
                    parts.append(delta)
                    yield _sse({"delta": delta})
        except Exception as e:  # noqa: BLE001
            log.error("OpenAI 호환 스트리밍 오류(%s): %s", AI_BASE_URL, e)
            yield _sse({"error": "AI 응답 생성 중 오류가 발생했습니다."})
            return
        answer = "".join(parts).strip()
        question = req.messages[-1].content
        cited_refs = " ".join(re.findall(r"「([^」]{2,40})」", answer))
        excerpts = _search_excerpts(f"{question} {cited_refs}")
        yield _sse({
            "citations": excerpts,
            "citations_label": "📄 관련 규정 원문 (자동 검색 — 답변 검증용)",
        })
        yield "data: [DONE]\n\n"

    return StreamingResponse(
        gen(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.post("/api/chat/stream")
def chat_stream(req: ChatRequest, x_app_password: str = Header(default="")):
    _check_password(x_app_password)
    if not policy_documents:
        raise HTTPException(503, "규정 문서가 아직 준비되지 않았습니다.")
    if not req.messages or req.messages[-1].role != "user":
        raise HTTPException(400, "마지막 메시지는 user 여야 합니다.")

    if MODE == "openai":
        return _openai_chat_stream(req)

    # 스트리밍 미지원 모드(claude/demo): 기존 로직으로 만든 뒤 한 번에 흘려보낸다.
    data = chat(req, x_app_password)

    def once():
        for seg in data.get("segments", []):
            if seg.get("text"):
                yield _sse({"delta": seg["text"]})
        yield _sse({
            "citations": data.get("citations", []),
            "citations_label": data.get("citations_label", ""),
        })
        yield "data: [DONE]\n\n"

    return StreamingResponse(
        once(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.get("/api/health")
def health(x_app_password: str = Header(default="")):
    info = {"status": "ok", "mode": MODE}
    # 문서 목록 등 상세 정보는 비밀번호 확인 후에만 노출
    if not APP_PASSWORD or hmac.compare_digest(x_app_password or "", APP_PASSWORD):
        info["model"] = AI_MODEL if MODE == "openai" else (MODEL if MODE == "claude" else None)
        info["documents"] = [d["title"] for d in policy_documents]
    return info


@app.get("/")
def index():
    return FileResponse(BASE_DIR / "static" / "index.html")


app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")


# ── NVIDIA 임베딩 프록시 (RAGFlow용) ─────────────────────────────────
# RAGFlow의 OpenAI 호환 클라이언트가 비표준 필드(drop_params)를 보내
# NVIDIA가 400으로 거부한다. 허용 필드만 남겨 전달하는 얇은 프록시.
# 경로에 토큰을 넣어 외부 오남용을 막는다 (EMBED_PROXY_TOKEN).
import httpx
from fastapi import Request
from fastapi.responses import JSONResponse

NVIDIA_API_KEY = os.getenv("NVIDIA_API_KEY", "")
EMBED_PROXY_TOKEN = os.getenv("EMBED_PROXY_TOKEN", "")
_EMBED_ALLOWED = {"input", "model", "user", "input_type", "encoding_format",
                  "dimensions", "embedding_type", "truncate", "modality"}


@app.post("/api/embedproxy/{token}/v1/embeddings")
async def embed_proxy(token: str, request: Request):
    if not EMBED_PROXY_TOKEN or token != EMBED_PROXY_TOKEN or not NVIDIA_API_KEY:
        raise HTTPException(status_code=404)
    body = await request.json()
    clean = {k: v for k, v in body.items() if k in _EMBED_ALLOWED}
    clean.setdefault("truncate", "END")
    async with httpx.AsyncClient(timeout=90) as client:
        r = await client.post(
            "https://integrate.api.nvidia.com/v1/embeddings",
            headers={"Authorization": "Bearer " + NVIDIA_API_KEY},
            json=clean,
        )
    try:
        payload = r.json()
    except ValueError:
        payload = {"error": r.text[:300]}
    return JSONResponse(status_code=r.status_code, content=payload)
