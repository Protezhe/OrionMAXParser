#!/usr/bin/env python3
import argparse
import json
import re
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path


TIME_RE = re.compile(r"\b([01]?\d|2[0-3]):([0-5]\d)\b")


def maybe_fix_mojibake(text: str) -> str:
    # Typical case: UTF-8 text decoded as CP1251.
    try:
        return text.encode("cp1251", errors="strict").decode("utf-8", errors="strict")
    except UnicodeError:
        return text


def normalize_hhmm(value: str) -> str | None:
    match = TIME_RE.fullmatch(value.strip())
    if not match:
        return None
    return f"{int(match.group(1)):02d}:{match.group(2)}"


def normalize_message_date(value: str) -> str | None:
    raw = value.strip()
    if not raw:
        return None

    if len(raw) == 10:
        try:
            return datetime.strptime(raw, "%Y-%m-%d").strftime("%Y-%m-%d")
        except ValueError:
            return None

    if raw.endswith("Z"):
        raw = raw[:-1] + "+00:00"
    try:
        dt = datetime.fromisoformat(raw)
    except ValueError:
        return None
    return dt.strftime("%Y-%m-%d")


def _post_json(url: str, payload: dict, timeout: float) -> dict:
    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        body = resp.read().decode("utf-8")
    return json.loads(body)


def detect_first_model(endpoint: str, timeout: float) -> str:
    tags_url = endpoint.rstrip("/").replace("/api/generate", "/api/tags")
    req = urllib.request.Request(tags_url, method="GET")
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    models = data.get("models") or []
    if not models:
        raise RuntimeError("No models found in Ollama (/api/tags returned empty list).")
    first = models[0].get("name")
    if not first:
        raise RuntimeError("Failed to detect model name from /api/tags.")
    return str(first)


def call_ollama(model: str, prompt: str, endpoint: str, timeout: float, retries: int) -> str:
    generate_payload = {
        "model": model,
        "prompt": prompt,
        "stream": False,
        "format": "json",
        "options": {"temperature": 0},
    }
    last_error: Exception | None = None
    for attempt in range(retries + 1):
        try:
            data = _post_json(endpoint, generate_payload, timeout=timeout)
            return str(data.get("response", "")).strip()
        except urllib.error.HTTPError as exc:
            if exc.code != 404:
                last_error = exc
                if attempt < retries:
                    time.sleep(0.5)
                    continue
                raise RuntimeError(f"Ollama request failed: {exc}") from exc
            break
        except urllib.error.URLError as exc:
            last_error = exc
            if attempt < retries:
                time.sleep(0.5)
                continue
            raise RuntimeError(f"Ollama request failed: {exc}") from exc

    # Fallback for servers that expose only /api/chat.
    chat_endpoint = endpoint.rstrip("/").replace("/api/generate", "/api/chat")
    chat_payload = {
        "model": model,
        "messages": [{"role": "user", "content": prompt}],
        "stream": False,
        "format": "json",
        "options": {"temperature": 0},
    }
    for attempt in range(retries + 1):
        try:
            data = _post_json(chat_endpoint, chat_payload, timeout=timeout)
            message = data.get("message", {})
            return str(message.get("content", "")).strip()
        except urllib.error.URLError as exc:
            last_error = exc
            if attempt < retries:
                time.sleep(0.5)
                continue
            raise RuntimeError(f"Ollama request failed: {exc}") from exc

    raise RuntimeError(f"Ollama request failed: {last_error}")


def build_prompt(text: str, message_date: str, message_time: str) -> str:
    return (
        "Ты извлекаешь событие из одного сообщения дежурного чата.\n"
        "Верни ТОЛЬКО JSON-объект БЕЗ комментариев со схемой:\n"
        "{\n"
        '  "category": "status|attendance|other",\n'
        '  "pavilion": "строка или null",\n'
        '  "status": "open|closed|repair|open_after_repair|null",\n'
        '  "attendance_count": "целое число или null",\n'
        '  "session_time": "HH:MM или null",\n'
        '  "time_from_text": "HH:MM или null"\n'
        "}\n"
        "Правила:\n"
        "- status=repair, если явно про ремонт/покрасочные/техработы.\n"
        "- status=open_after_repair, если явно после ремонта снова работает.\n"
        "- attendance_count только если это сообщение про посетителей/билеты на сеанс.\n"
        "- session_time это время сеанса для attendance.\n"
        "- time_from_text это время события, если оно есть в тексте.\n"
        f"messageDate: {message_date}\n"
        f"messageTime: {message_time}\n"
        f"message: {text}"
    )


def normalize_model_result(raw: str) -> dict:
    cleaned = raw.strip()
    if cleaned.startswith("```"):
        lines = cleaned.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        cleaned = "\n".join(lines).strip()

    if cleaned and cleaned[0] != "{":
        start = cleaned.find("{")
        end = cleaned.rfind("}")
        if start != -1 and end != -1 and end > start:
            cleaned = cleaned[start : end + 1]

    try:
        obj = json.loads(cleaned)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Ollama returned non-JSON: {raw}") from exc

    category = obj.get("category")
    if category not in {"status", "attendance", "other"}:
        category = "other"

    status = obj.get("status")
    if status not in {"open", "closed", "repair", "open_after_repair", None}:
        status = None

    count = obj.get("attendance_count")
    if isinstance(count, str) and count.isdigit():
        count = int(count)
    if not isinstance(count, int):
        count = None

    session_time = obj.get("session_time")
    if session_time is not None:
        session_time = normalize_hhmm(str(session_time).strip())

    time_from_text = obj.get("time_from_text")
    if time_from_text is not None:
        time_from_text = normalize_hhmm(str(time_from_text).strip())

    return {
        "category": category,
        "pavilion": obj.get("pavilion"),
        "status": status,
        "attendance_count": count,
        "session_time": session_time,
        "time_from_text": time_from_text,
    }


def map_status_to_ru(status: str | None) -> str | None:
    mapping = {
        "open": "открыт",
        "closed": "закрыт",
        "repair": "на ремонте",
        "open_after_repair": "открыт с ремонта",
    }
    return mapping.get(status)


def choose_event_time(model_time: str | None, raw_text: str, fallback_time: str) -> str:
    if isinstance(model_time, str):
        normalized = normalize_hhmm(model_time)
        if normalized:
            return normalized

    match = TIME_RE.search(raw_text)
    if match:
        return f"{int(match.group(1)):02d}:{match.group(2)}"

    return fallback_time


def enrich_message(message: dict, model: str, endpoint: str, timeout: float, retries: int) -> dict:
    text_raw = str(message.get("text", "")).strip()
    message_date = normalize_message_date(str(message.get("messageDate", "")))
    if not message_date:
        raise RuntimeError("messageDate is missing or invalid")

    message_time = normalize_hhmm(str(message.get("time", "")).strip()) or "00:00"
    text = maybe_fix_mojibake(text_raw)

    prompt = build_prompt(text=text, message_date=message_date, message_time=message_time)
    raw = call_ollama(model=model, prompt=prompt, endpoint=endpoint, timeout=timeout, retries=retries)
    parsed = normalize_model_result(raw)

    event_time = choose_event_time(parsed.get("time_from_text"), text, fallback_time=message_time)

    out = dict(message)
    out["Павильон"] = parsed.get("pavilion")
    out["Статус"] = map_status_to_ru(parsed.get("status"))
    out["Дата"] = message_date
    out["Время"] = event_time
    out["ИсточникВремени"] = "text" if parsed.get("time_from_text") or TIME_RE.search(text) else "message_time"
    out["ПосетителейНаСеанс"] = parsed.get("attendance_count")
    out["ВремяСеанса"] = parsed.get("session_time")
    out["КатегорияСобытия"] = parsed.get("category")
    return out


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extract pavilion status and attendance events from parse.json using Ollama"
    )
    parser.add_argument("-i", "--input", default="parse.json", help="Input JSON path")
    parser.add_argument("-o", "--output", default="parse.enriched.json", help="Output JSON path")
    parser.add_argument(
        "--model",
        default="auto",
        help="Ollama model name (default: auto, first model from /api/tags)",
    )
    parser.add_argument(
        "--endpoint",
        default="http://127.0.0.1:11434/api/generate",
        help="Ollama generate endpoint",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=0,
        help="How many first messages to process (0 = all messages)",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        default=20.0,
        help="Timeout in seconds for each Ollama HTTP request (default: 20)",
    )
    parser.add_argument(
        "--retries",
        type=int,
        default=1,
        help="How many retries for failed Ollama requests (default: 1)",
    )
    parser.add_argument(
        "--fail-fast",
        action="store_true",
        help="Stop immediately on first Ollama error",
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)

    with input_path.open("r", encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, dict):
        raise RuntimeError("Input file must contain a JSON object with a messages array")

    messages = data.get("messages")
    if not isinstance(messages, list):
        raise RuntimeError("Input JSON must contain messages as a list")

    total = len(messages)
    limit = total if args.limit <= 0 else min(args.limit, total)

    model = args.model
    if limit > 0 and model == "auto":
        model = detect_first_model(args.endpoint, timeout=args.timeout)
        print(f"Auto model selected: {model}", file=sys.stderr)

    for idx in range(limit):
        try:
            messages[idx] = enrich_message(
                messages[idx],
                model=model,
                endpoint=args.endpoint,
                timeout=args.timeout,
                retries=max(args.retries, 0),
            )
            print(f"Processed {idx + 1}/{limit}", file=sys.stderr)
        except Exception as exc:
            print(f"Error on message {idx + 1}/{limit}: {exc}", file=sys.stderr)
            if args.fail_fast:
                raise
            messages[idx] = dict(messages[idx])
            messages[idx]["parser_error"] = str(exc)

    data["messages"] = messages
    data["count"] = len(messages)

    with output_path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"Input: {input_path}")
    print(f"Output: {output_path}")
    print(f"Processed messages: {limit}")
    print(f"Model: {model}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
