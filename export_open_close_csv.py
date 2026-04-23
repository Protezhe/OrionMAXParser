import argparse
import csv
import json
from collections import defaultdict
from pathlib import Path


STATUS_OPEN = {"открыт", "open", "opened"}
STATUS_CLOSE = {"закрыт", "closed", "close"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Экспортирует CSV (дата, аттракцион, время открытия, время закрытия) "
            "из parse_dedup.enriched.json."
        )
    )
    parser.add_argument(
        "--input",
        default="parse_dedup.enriched.json",
        help="Путь к входному JSON файлу.",
    )
    parser.add_argument(
        "--output",
        default="attractions_open_close.csv",
        help="Путь к выходному CSV файлу.",
    )
    return parser.parse_args()


def normalize_status(value: object) -> str | None:
    if value is None:
        return None
    status = str(value).strip().lower()
    if status in STATUS_OPEN:
        return "open"
    if status in STATUS_CLOSE:
        return "close"
    return None


def main() -> None:
    args = parse_args()
    input_path = Path(args.input)
    output_path = Path(args.output)

    raw = json.loads(input_path.read_text(encoding="utf-8"))
    messages = raw.get("messages", [])

    grouped: dict[tuple[str, str], dict[str, str]] = defaultdict(
        lambda: {"date": "", "name": "", "open_time": "", "close_time": ""}
    )

    skipped = 0
    used_events = 0

    for item in messages:
        if not isinstance(item, dict):
            skipped += 1
            continue

        name = item.get("Павильон")
        date = item.get("Дата")
        time = item.get("Время")
        status = normalize_status(item.get("Статус"))

        # По требованию используем только структурные поля.
        if not all([name, date, time, status]):
            skipped += 1
            continue

        key = (str(date), str(name))
        row = grouped[key]
        row["date"] = str(date)
        row["name"] = str(name)

        if status == "open":
            current = row["open_time"]
            row["open_time"] = min(current, str(time)) if current else str(time)
        elif status == "close":
            current = row["close_time"]
            row["close_time"] = max(current, str(time)) if current else str(time)

        used_events += 1

    rows = sorted(grouped.values(), key=lambda x: (x["date"], x["name"]))

    with output_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["дата", "название аттракциона", "время открытия", "время закрытия"])
        for row in rows:
            writer.writerow([row["date"], row["name"], row["open_time"], row["close_time"]])

    print(f"Готово. CSV: {output_path}")
    print(f"Использовано событий: {used_events}")
    print(f"Пропущено сообщений: {skipped}")
    print(f"Итоговых строк: {len(rows)}")


if __name__ == "__main__":
    main()
