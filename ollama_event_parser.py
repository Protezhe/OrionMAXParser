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
OUTPUT_FIELDS = (
    "text",
    "Павильон",
    "Статус",
    "Дата",
    "Время",
    "ИсточникВремени",
    "ПосетителейНаСеанс",
    "ВремяСеанса",
    "КатегорияСобытия",
)


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


def _clean_json_text(raw: str) -> str:
    cleaned = raw.strip()
    if cleaned.startswith("```"):
        lines = cleaned.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        cleaned = "\n".join(lines).strip()
    return cleaned


def build_batch_prompt(items: list[dict]) -> str:
    payload = json.dumps(items, ensure_ascii=False)
    return (
        "Ты извлекаешь события из сообщений дежурного чата.\n"
        "Верни ТОЛЬКО JSON-объект БЕЗ комментариев в формате:\n"
        "{\n"
        '  "items": [\n'
        "    {\n"
        '      "id": 0,\n'
        '      "category": "status|attendance|other",\n'
        '      "pavilion": "строка или null",\n'
        '      "status": "open|closed|repair|open_after_repair|null",\n'
        '      "attendance_count": "целое число или null",\n'
        '      "session_time": "HH:MM или null",\n'
        '      "time_from_text": "HH:MM или null"\n'
        "    }\n"
        "  ]\n"
        "}\n"
        "Правила:\n"
        "- Для каждого входного элемента должен быть ровно один выходной элемент с тем же id.\n"
        "- status=repair, если явно про ремонт/покрасочные/техработы.\n"
        "- status=open_after_repair, если явно после ремонта снова работает.\n"
        "- attendance_count только если это сообщение про посетителей/билеты на сеанс.\n"
        "- session_time это время сеанса для attendance.\n"
        "- time_from_text это время события, если оно есть в тексте.\n"
        f"messages: {payload}"
    )


def normalize_model_item(obj: dict) -> dict:
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


def normalize_batch_result(raw: str, expected_size: int) -> list[dict]:
    cleaned = _clean_json_text(raw)

    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError:
        start_obj = cleaned.find("{")
        end_obj = cleaned.rfind("}")
        start_arr = cleaned.find("[")
        end_arr = cleaned.rfind("]")
        if start_obj != -1 and end_obj > start_obj:
            data = json.loads(cleaned[start_obj : end_obj + 1])
        elif start_arr != -1 and end_arr > start_arr:
            data = json.loads(cleaned[start_arr : end_arr + 1])
        else:
            raise RuntimeError(f"Ollama returned non-JSON: {raw}")

    if isinstance(data, dict):
        items = data.get("items")
    elif isinstance(data, list):
        items = data
    else:
        raise RuntimeError(f"Unexpected model response type: {type(data).__name__}")

    if not isinstance(items, list):
        raise RuntimeError("Model response must contain list in 'items'")
    if len(items) != expected_size:
        raise RuntimeError(f"Batch size mismatch: expected {expected_size}, got {len(items)}")

    by_id: dict[int, dict] = {}
    can_use_ids = True

    for item in items:
        if not isinstance(item, dict):
            raise RuntimeError("Each batch item from model must be a JSON object")
        item_id = item.get("id")
        if isinstance(item_id, int) and 0 <= item_id < expected_size and item_id not in by_id:
            by_id[item_id] = normalize_model_item(item)
        else:
            can_use_ids = False

    if can_use_ids and len(by_id) == expected_size:
        return [by_id[idx] for idx in range(expected_size)]

    return [normalize_model_item(item) for item in items]


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


def build_output_row(
    *,
    parsed: dict,
    message_date: str,
    message_time: str,
    source_text_raw: str,
    source_text: str,
) -> dict:
    event_time = choose_event_time(parsed.get("time_from_text"), source_text, fallback_time=message_time)
    return {
        "text": source_text_raw,
        "Павильон": parsed.get("pavilion"),
        "Статус": map_status_to_ru(parsed.get("status")),
        "Дата": message_date,
        "Время": event_time,
        "ИсточникВремени": "text" if parsed.get("time_from_text") or TIME_RE.search(source_text) else "message_time",
        "ПосетителейНаСеанс": parsed.get("attendance_count"),
        "ВремяСеанса": parsed.get("session_time"),
        "КатегорияСобытия": parsed.get("category"),
    }


def build_fallback_row(message: dict) -> dict:
    message_date = normalize_message_date(str(message.get("messageDate", ""))) or ""
    message_time = normalize_hhmm(str(message.get("time", "")).strip()) or "00:00"
    source_text = str(message.get("text", "")).strip()
    return {
        "text": maybe_fix_mojibake(source_text),
        "Павильон": None,
        "Статус": None,
        "Дата": message_date,
        "Время": message_time,
        "ИсточникВремени": "message_time",
        "ПосетителейНаСеанс": None,
        "ВремяСеанса": None,
        "КатегорияСобытия": "other",
    }


def compact_output_row(row: dict) -> dict:
    return {field: row.get(field) for field in OUTPUT_FIELDS}


def load_existing_output_messages(output_path: Path) -> list[dict]:
    if not output_path.exists():
        return []

    with output_path.open("r", encoding="utf-8") as f:
        raw = json.load(f)

    if isinstance(raw, dict):
        messages = raw.get("messages", [])
    elif isinstance(raw, list):
        messages = raw
    else:
        raise RuntimeError("Output JSON must be an object with messages or a list")

    if not isinstance(messages, list):
        raise RuntimeError("Output JSON field 'messages' must be a list")

    normalized = []
    for item in messages:
        if isinstance(item, dict):
            normalized.append(compact_output_row(item))
    return normalized


def save_output(
    *,
    output_path: Path,
    input_path: Path,
    source_count: int,
    processed_now: int,
    model: str,
    messages: list[dict],
) -> None:
    payload = {
        "input": str(input_path),
        "sourceCount": source_count,
        "count": len(messages),
        "processedThisRun": processed_now,
        "model": model,
        "messages": messages,
    }
    with output_path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def enrich_batch(messages_batch: list[dict], model: str, endpoint: str, timeout: float, retries: int) -> list[dict]:
    prepared = []
    prompt_items = []
    for idx, message in enumerate(messages_batch):
        text_raw = str(message.get("text", "")).strip()
        message_date = normalize_message_date(str(message.get("messageDate", "")))
        if not message_date:
            raise RuntimeError(f"messageDate is missing or invalid at batch item {idx}")
        message_time = normalize_hhmm(str(message.get("time", "")).strip()) or "00:00"
        text = maybe_fix_mojibake(text_raw)

        prepared.append({"message": message, "message_date": message_date, "message_time": message_time, "text": text})
        prompt_items.append({"id": idx, "messageDate": message_date, "messageTime": message_time, "message": text})

    prompt = build_batch_prompt(prompt_items)
    raw = call_ollama(model=model, prompt=prompt, endpoint=endpoint, timeout=timeout, retries=retries)
    parsed_items = normalize_batch_result(raw, expected_size=len(messages_batch))

    output = []
    for idx, parsed in enumerate(parsed_items):
        src = prepared[idx]
        out = build_output_row(
            parsed=parsed,
            message_date=src["message_date"],
            message_time=src["message_time"],
            source_text_raw=src["text"],
            source_text=src["text"],
        )
        output.append(out)

    return output


def enrich_batch_with_split(
    messages_batch: list[dict],
    model: str,
    endpoint: str,
    timeout: float,
    retries: int,
    fail_fast: bool,
) -> list[dict]:
    try:
        return enrich_batch(
            messages_batch=messages_batch,
            model=model,
            endpoint=endpoint,
            timeout=timeout,
            retries=retries,
        )
    except Exception as exc:
        size = len(messages_batch)
        if size <= 1:
            if fail_fast:
                raise
            print(f"Single message failed, writing fallback row: {exc}", file=sys.stderr)
            return [build_fallback_row(messages_batch[0])]

        mid = size // 2
        print(
            f"Batch of {size} failed ({exc}). Retrying as {mid}+{size - mid}.",
            file=sys.stderr,
        )
        left = enrich_batch_with_split(
            messages_batch=messages_batch[:mid],
            model=model,
            endpoint=endpoint,
            timeout=timeout,
            retries=retries,
            fail_fast=fail_fast,
        )
        right = enrich_batch_with_split(
            messages_batch=messages_batch[mid:],
            model=model,
            endpoint=endpoint,
            timeout=timeout,
            retries=retries,
            fail_fast=fail_fast,
        )
        return left + right


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
        "--batch-size",
        type=int,
        default=10,
        help="How many messages to send to model in one request (default: 10)",
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

    source_messages = data.get("messages")
    if not isinstance(source_messages, list):
        raise RuntimeError("Input JSON must contain messages as a list")

    source_total = len(source_messages)
    ready_messages = load_existing_output_messages(output_path)
    start_index = len(ready_messages)
    if start_index > source_total:
        raise RuntimeError(
            f"Output has more rows than input messages ({start_index} > {source_total}). "
            "Delete output file or use another --output path."
        )

    end_index = source_total if args.limit <= 0 else min(source_total, start_index + args.limit)
    run_total = max(0, end_index - start_index)

    if not output_path.exists():
        save_output(
            output_path=output_path,
            input_path=input_path,
            source_count=source_total,
            processed_now=0,
            model=args.model,
            messages=ready_messages,
        )

    model = args.model
    if run_total > 0 and model == "auto":
        model = detect_first_model(args.endpoint, timeout=args.timeout)
        print(f"Auto model selected: {model}", file=sys.stderr)

    batch_size = max(1, args.batch_size)
    processed_now = 0

    for start in range(start_index, end_index, batch_size):
        end = min(start + batch_size, end_index)
        try:
            enriched = enrich_batch_with_split(
                source_messages[start:end],
                model=model,
                endpoint=args.endpoint,
                timeout=args.timeout,
                retries=max(args.retries, 0),
                fail_fast=args.fail_fast,
            )
            ready_messages.extend(enriched)
            processed_now += len(enriched)
            save_output(
                output_path=output_path,
                input_path=input_path,
                source_count=source_total,
                processed_now=processed_now,
                model=model,
                messages=ready_messages,
            )
            print(f"Processed {end}/{end_index}", file=sys.stderr)
        except Exception as exc:
            print(f"Error on batch {start + 1}-{end}/{end_index}: {exc}", file=sys.stderr)
            if args.fail_fast:
                raise
            fallback_rows = [build_fallback_row(msg) for msg in source_messages[start:end]]
            ready_messages.extend(fallback_rows)
            processed_now += len(fallback_rows)
            save_output(
                output_path=output_path,
                input_path=input_path,
                source_count=source_total,
                processed_now=processed_now,
                model=model,
                messages=ready_messages,
            )

    print(f"Input: {input_path}")
    print(f"Output: {output_path}")
    print(f"Start index: {start_index}")
    print(f"Processed this run: {processed_now}")
    print(f"Total in output: {len(ready_messages)}")
    print(f"Model: {model}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
