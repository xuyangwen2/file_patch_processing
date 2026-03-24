# -*- coding: utf-8 -*-
import os
import json
import uuid
import time
from typing import Any, Dict, List

import pytest
import requests

def getenv_str(name: str, default: str = "") -> str:
    return os.getenv(name, default).strip()


BASE_URL = getenv_str("BASE_URL", "https://dev-ssc.antagent.space/api").rstrip("/")
AUTH_TOKEN = getenv_str(
    "AUTH_TOKEN",
    "",
)
TIMEOUT = float(getenv_str("TIMEOUT", "20"))

HEADERS = {"Accept": "application/json"}
if AUTH_TOKEN:
    HEADERS["Authorization"] = f"Bearer {AUTH_TOKEN}"


def url(path: str) -> str:
    return f"{BASE_URL}{path}"


def do_get(path: str, params: Dict[str, Any]):
    t0 = time.time()
    r = requests.get(url(path), params=params, headers=HEADERS, timeout=TIMEOUT)
    return r, int((time.time() - t0) * 1000)


def do_post_json(path: str, body: Dict[str, Any], *, stream: bool = False):
    h = dict(HEADERS)
    h["Content-Type"] = "application/json"
    t0 = time.time()
    r = requests.post(url(path), json=body, headers=h, timeout=TIMEOUT, stream=stream)
    return r, int((time.time() - t0) * 1000)


def must_json(resp: requests.Response) -> Any:
    try:
        return resp.json()
    except Exception:
        pytest.fail(f"Response is not JSON. status={resp.status_code}, text={resp.text[:500]}")


def assert_base_envelope(body: Dict[str, Any]):
    assert isinstance(body, dict)
    assert body.get("status") == 0, f"status!=0: {body}"
    assert body.get("message") == "请求成功", f"message!=请求成功: {body.get('message')}"
    assert "result" in body
    assert "timestamp" in body


def assert_paged_result(result: Dict[str, Any], page_num: int, page_size: int):
    assert isinstance(result, dict)
    for k in ["pageNum", "pageSize", "size", "pages", "total", "list"]:
        assert k in result, f"result missing {k}"
    assert str(result["pageNum"]) == str(page_num)
    assert str(result["pageSize"]) == str(page_size)

    assert isinstance(result["size"], int) and result["size"] >= 0
    assert isinstance(result["pages"], int) and result["pages"] >= 0
    assert isinstance(result["total"], int) and result["total"] >= 0
    assert isinstance(result["list"], list)
    assert result["size"] == len(result["list"]), "size 必须等于 len(list)"

    if result["total"] == 0:
        assert result["pages"] == 0, "total==0 时 pages 期望为 0"


def is_highlighted(s: str) -> bool:
    s2 = s.lower()
    return ("<em>" in s2 and "</em>" in s2) or ("<mark>" in s2 and "</mark>" in s2)


def dedup_ids(items: List[Dict[str, Any]]) -> List[str]:
    ids = []
    for it in items:
        if isinstance(it, dict) and "id" in it and it["id"] is not None:
            ids.append(str(it["id"]))
    return ids


def _show(label: str, data: Any, cost_ms: int = 0) -> None:
    print(f"\n{'='*60}")
    print(f"[{label}]" + (f"  耗时 {cost_ms}ms" if cost_ms else ""))
    print(json.dumps(data, ensure_ascii=False, indent=2))
    print('='*60)


# ========== 用例参数 ==========
PAGE_NUM = "1"
PAGE_SIZE = "3"

CARDS_PARAMS = {
    "query": os.getenv("CARDS_QUERY", "你的模型是？"),
    "pageNum": PAGE_NUM,
    "pageSize": PAGE_SIZE,
    "requestId": os.getenv("REQUEST_ID", str(uuid.uuid4())),
}

DOCS_PARAMS = {
    "query": os.getenv("DOCS_QUERY", "考勤规则是什么?"),
    "pageNum": PAGE_NUM,
    "pageSize": PAGE_SIZE,
    "requestId": os.getenv("REQUEST_ID", str(uuid.uuid4())),
    "sourceFilter": os.getenv("DOCS_SOURCE_FILTER", "pota,ant"),
}

RAG_PARAMS = {
    "query": os.getenv("RAG_QUERY", "精益求精"),
    "sourceFilter": os.getenv("RAG_SOURCE_FILTER", "pota,appia"),
}

RAG_SHORT_PARAMS = {
    "query": os.getenv("RAG_QUERY", "精益求精"),
    "pageNum": "1",
    "pageSize": "2",
}

RAG_FULL_PARAMS = {
    "query": os.getenv("RAG_QUERY", "精益求精"),
    "pageNum": "1",
    "pageSize": "2",
}

RAG_REWRITE_PARAMS = {"query": os.getenv("RAG_QUERY", "精益求精")}

CHAT_BODY = {"messages": [{"role": "user", "content": os.getenv("CHAT_QUESTION", "你好")}]}
CHAT_TIMEOUT = float(getenv_str("CHAT_TIMEOUT", "60"))

RAG_BULK_BODY = {
    "esQueries": [
        {"query": "1684", "size": 10, "fragmentSize": 400, "fragmentNum": 10},
        {"query": "POTA 挑战度 级别", "size": 10, "fragmentSize": 400, "fragmentNum": 10},
        {"query": "POTA 挑战度", "size": 10, "fragmentSize": 600, "fragmentNum": 10},
        {"query": "POTA 挑战度有几档？", "size": 10, "fragmentSize": 600, "fragmentNum": 10},
    ],
    "kgQueries": [{"query": "吴招乐"}],
}


# ========== 业务结构断言 ==========
def assert_docs_item(it: Dict[str, Any]):
    for k in [
        "id", "title", "content", "owner", "ownerId", "ownerName",
        "source", "sourceCode", "type", "ownerInfo",
        "outline", "summary", "score", "updatedAt"
    ]:
        assert k in it, f"docs item missing {k}"
    assert isinstance(it["id"], (str, int))
    assert isinstance(it["title"], str) and it["title"]
    assert isinstance(it["content"], str)
    assert isinstance(it["ownerId"], int)
    assert isinstance(it["score"], (int, float))

    oi = it["ownerInfo"]
    if oi is not None:
        assert isinstance(oi, dict)
        for k in ["id", "username", "name", "email", "jobList"]:
            assert k in oi


def assert_rag_item(it: Dict[str, Any]):
    for k in ["id", "title", "content", "type", "url", "sourceCode", "summary", "outline", "queries"]:
        assert k in it, f"rag item missing {k}"
    assert isinstance(it["id"], int)
    assert isinstance(it["title"], str)
    assert isinstance(it["content"], str)
    assert isinstance(it["type"], str)
    assert (it["url"] is None) or isinstance(it["url"], str)
    assert (it["sourceCode"] is None) or isinstance(it["sourceCode"], str)
    assert (it["queries"] is None) or isinstance(it["queries"], list)


def parse_streaming_json_lines(resp: requests.Response, max_events: int = 2000) -> List[Dict[str, Any]]:
    """
    兼容：
    - SSE: data: {...}
    - JSON Lines: {...}\n{...}
    """
    events: List[Dict[str, Any]] = []
    buf = ""
    for raw in resp.iter_lines(decode_unicode=True):
        if raw is None:
            continue
        line = raw.strip()
        if not line:
            continue

        if line.startswith("data:"):
            payload = line[len("data:"):].strip()
            if payload in ("[DONE]", "DONE"):
                break
            try:
                events.append(json.loads(payload))
            except Exception:
                buf += payload
                try:
                    events.append(json.loads(buf))
                    buf = ""
                except Exception:
                    pass
        else:
            try:
                events.append(json.loads(line))
            except Exception:
                buf += line
                try:
                    events.append(json.loads(buf))
                    buf = ""
                except Exception:
                    pass

        if len(events) >= max_events:
            break

    return events


def assert_chat_stream_events(events: List[Dict[str, Any]]) -> str:
    assert events, "chat 必须返回至少 1 个流式事件"
    for e in events[:50]:
        assert "content" in e
        assert "type" in e
        assert (e["content"] is None) or isinstance(e["content"], str)
        assert (e["type"] is None) or isinstance(e["type"], str)
        if "timestamp" in e:
            assert (e["timestamp"] is None) or isinstance(e["timestamp"], str)

    combined = "".join([e.get("content") or "" for e in events]).strip()
    assert combined, "拼接后的 content 不能为空"
    return combined

# ========== cards ==========
def test_search_cards_schema_pagination():
    resp, cost = do_get("/search/cards", CARDS_PARAMS)
    assert resp.status_code == 200, resp.text[:300]
    body = must_json(resp)

    assert_base_envelope(body)
    assert_paged_result(body["result"], CARDS_PARAMS["pageNum"], CARDS_PARAMS["pageSize"])
    assert cost < 8000
    _show("cards", body["result"], cost)


# ========== docs ==========
def test_search_docs_schema_highlight_dedup_sourcefilter():
    resp, cost = do_get("/search/docs", DOCS_PARAMS)
    assert resp.status_code == 200, resp.text[:300]
    body = must_json(resp)

    assert_base_envelope(body)
    assert_paged_result(body["result"], DOCS_PARAMS["pageNum"], DOCS_PARAMS["pageSize"])

    items = body["result"]["list"]
    for it in items:
        assert_docs_item(it)

    # 高亮：docs 的 title/content 至少一处包含 <em>..</em>（你示例就是这样）
    if items:
        merged = "\n".join([(it.get("title") or "") + "\n" + (it.get("content") or "") for it in items])
        assert is_highlighted(merged), "docs 未检测到高亮 <em>/<mark>"

    # 去重：同一页 id 不允许重复
    ids = dedup_ids(items)
    assert len(ids) == len(set(ids)), f"docs 去重失败：{ids}"

    # sourceFilter 生效：如果请求带了 sourceFilter，则返回的 sourceCode 应属于集合（若你们允许空则放宽）
    allowed = {s.strip() for s in DOCS_PARAMS.get("sourceFilter", "").split(",") if s.strip()}
    if allowed and items:
        for it in items:
            assert it.get("sourceCode") in allowed, f"sourceFilter={allowed} 但返回 sourceCode={it.get('sourceCode')}"

    assert cost < 12000
    _show("docs", body["result"], cost)


# ========== rag ==========
def test_rag_schema_and_sourcefilter():
    resp, cost = do_get("/search/chat/rag", RAG_PARAMS)
    assert resp.status_code == 200, resp.text[:300]
    body = must_json(resp)

    assert_base_envelope(body)
    assert isinstance(body["result"], list)
    for it in body["result"]:
        assert_rag_item(it)

    # sourceFilter 生效（同上逻辑）
    allowed = {s.strip() for s in RAG_PARAMS.get("sourceFilter", "").split(",") if s.strip()}
    if allowed and body["result"]:
        for it in body["result"]:
            sc = it.get("sourceCode")
            # url/sourceCode 允许为 null，但如果非 null 必须在 filter 内
            if sc is not None:
                assert sc in allowed, f"sourceFilter={allowed} 但返回 sourceCode={sc}"

    assert cost < 12000
    _show("rag", body["result"], cost)


def test_rag_short_schema_and_paging_params_reflected():
    resp, _ = do_get("/search/chat/rag-short", RAG_SHORT_PARAMS)
    assert resp.status_code == 200, resp.text[:300]
    body = must_json(resp)

    assert_base_envelope(body)
    assert isinstance(body["result"], list)
    for it in body["result"]:
        assert_rag_item(it)

    # 你给的参数是 pageSize=2；做一个“上限校验”（不强制等于2，但不应超过2）
    assert len(body["result"]) <= int(RAG_SHORT_PARAMS["pageSize"])
    _show("rag-short", body["result"])


def test_rag_full_schema_and_paging_params_reflected():
    resp, _ = do_get("/search/chat/rag-full", RAG_FULL_PARAMS)
    assert resp.status_code == 200, resp.text[:300]
    body = must_json(resp)

    assert_base_envelope(body)
    assert isinstance(body["result"], list)
    for it in body["result"]:
        assert_rag_item(it)
    assert len(body["result"]) <= int(RAG_FULL_PARAMS["pageSize"])
    _show("rag-full", body["result"])


def test_rag_rewrite_schema():
    resp, _ = do_get("/search/chat/rag-rewrite", RAG_REWRITE_PARAMS)
    assert resp.status_code == 200, resp.text[:300]
    body = must_json(resp)

    assert_base_envelope(body)
    assert isinstance(body["result"], list)
    for it in body["result"]:
        assert_rag_item(it)
    _show("rag-rewrite", body["result"])


# ========== rag-bulk ==========
def test_rag_bulk_schema_and_queries_mapping():
    resp, cost = do_post_json("/search/chat/rag-bulk", RAG_BULK_BODY)
    assert resp.status_code == 200, resp.text[:300]
    body = must_json(resp)

    assert_base_envelope(body)
    assert isinstance(body["result"], list)

    allowed_queries = {q["query"] for q in RAG_BULK_BODY["esQueries"]} | {q["query"] for q in RAG_BULK_BODY["kgQueries"]}

    for it in body["result"]:
        assert_rag_item(it)
        qs = it.get("queries")
        if qs is not None:
            for q in qs:
                assert q in allowed_queries, f"bulk 返回 queries 出现未知 query：{q}"

    assert cost < 20000
    _show("rag-bulk", body["result"], cost)


# ========== chat（流式） ==========
def test_chat_streaming_contract_and_semantics():
    h = dict(HEADERS)
    h["Content-Type"] = "application/json"
    t0 = time.time()
    resp = requests.post(url("/search/chat"), json=CHAT_BODY, headers=h, timeout=CHAT_TIMEOUT, stream=True)
    cost = int((time.time() - t0) * 1000)

    assert resp.status_code == 200, resp.text[:300]

    events = parse_streaming_json_lines(resp)
    combined = assert_chat_stream_events(events)

    assert combined, "拼接后的 content 不能为空"
    assert cost < int(CHAT_TIMEOUT * 1000)
    _show("chat", {"event_count": len(events), "combined": combined}, cost)
