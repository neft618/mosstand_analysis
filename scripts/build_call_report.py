from __future__ import annotations

import json
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.trendline import Trendline
from openpyxl.styles import Alignment, Font, PatternFill


ROOT = Path(__file__).resolve().parents[1]
INPUT_XLSX = ROOT / "input" / "звонки_менеджеры_март-05июня.xlsx"
INPUT_JSON_DIR = ROOT / "input" / "calls_json"
OUTPUT_DIR = ROOT / "output"
OUTPUT_XLSX = OUTPUT_DIR / "звонки_менеджеры_март-05июня_анализ.xlsx"

POSITIVE_STAGES = {
    "Оплачен",
    "Запущено в производство",
    "Счет отдан в оплату",
    "Передано в логистику",
    "Передано дизайнеру",
    "УПД передано на подписании",
    "Согласие получено",
    "КП получено клиентом",
    "Потребность выявлена",
    "Назначен Ответственный",
}

NEGATIVE_STAGES = {
    "Отказ ms",
    "Проиграли по цене",
    "Низкий бюджет",
    "Недозвон",
    "Отложено",
    "Не удалось обработать",
    "Отмена мероприятия",
    "Отсутствие продукции ms",
    "Просрочено",
}

STAGE_WEIGHTS = {
    "positive": 0.2,
    "neutral": 0.6,
    "negative": 1.0,
}

ISSUES = [
    (
        "price_objection",
        "Цена / бюджет",
        [
            "дорог",
            "цена",
            "стоим",
            "бюджет",
            "дешев",
            "скид",
            "прайс",
        ],
        "Рано переводить цену в ценность: показывать пакеты, якорить вариантами и не давать скидку без обмена на условия.",
    ),
    (
        "postpone",
        "Перенос / подумать",
        [
            "подумаю",
            "перезвон",
            "созвоним",
            "не актуаль",
            "отлож",
        ],
        "Фиксировать следующий контакт с датой и временем, иначе сделка зависает без ответственности.",
    ),
    (
        "compare",
        "Сравнение с конкурентами",
        [
            "конкурент",
            "сравн",
            "аналог",
            "другой вариант",
            "вариантов",
            "где дешевле",
        ],
        "Отстраивать ценность по срокам, комплектации, сервису и примерам, а не спорить только по цене.",
    ),
    (
        "qualification_gap",
        "Недоквалификация",
        [
            "размер",
            "тираж",
            "срок",
            "доставка",
            "монтаж",
            "адрес",
            "материал",
            "оплата",
            "формат",
            "бюджет",
            "кто принимает решение",
            "лпр",
        ],
        "Задавать обязательный чек-лист вопросов по задаче, срокам, ЛПР, доставке и бюджету до отправки КП.",
    ),
    (
        "no_next_step",
        "Нет следующего шага",
        [
            "отправлю",
            "скину",
            "пришлю",
            "согласуем",
            "оформим",
            "подберу",
            "выберем",
            "жду ответ",
            "до связи",
            "созвон",
        ],
        "Закрывать каждый звонок на конкретное действие менеджера и клиента, иначе теряется контроль над сделкой.",
    ),
    (
        "upsell_missed",
        "Упущенный апсейл",
        [
            "монтаж",
            "дизайн",
            "доставка",
            "сроч",
            "сборк",
            "установк",
            "макет",
        ],
        "Пакетировать допуслуги и предлагать комплекс: дизайн, доставка, монтаж, срочное производство.",
    ),
]

CORE_PRODUCT_PATTERNS = {
    "роллап": ["роллап", "ролл-ап", "ролик"],
    "стенд": ["стенд", "стенда", "стенды"],
    "пресс-волл": ["пресс-волл", "press wall"],
    "баннер": ["баннер"],
    "флаг": ["флаг", "виндер"],
    "лайтбокс": ["лайтбокс", "световой бокс", "световой короб"],
    "буквы": ["букв", "объемн"],
    "вывеска": ["вывеск"],
    "табличка": ["табличк"],
    "печать": ["печать", "полиграф", "листовк", "плакат", "наклейк"],
    "стойка": ["стойк", "промостойк", "поп-ап"],
    "павильон": ["павильон", "застройк"],
}

SERVICE_PATTERNS = [
    "доставк",
    "монтаж",
    "установк",
    "сборк",
    "дизайн",
    "макет",
    "срочн",
]


def normalize(text: str) -> str:
    text = text.lower().replace("ё", "е")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def stage_class(stage: str) -> str:
    if stage in POSITIVE_STAGES:
        return "positive"
    if stage in NEGATIVE_STAGES:
        return "negative"
    return "neutral"


def stage_weight(stage: str) -> float:
    return STAGE_WEIGHTS[stage_class(stage)]


def read_json_text(json_path: Path) -> str:
    with json_path.open("r", encoding="utf-8") as fh:
        payload = json.load(fh)
    return normalize(" ".join(segment.get("text", "") for segment in payload.get("segments", [])))


def first_hit(text: str, patterns: Iterable[str]) -> str:
    for pattern in patterns:
        if pattern in text:
            return pattern
    return ""


def extract_product_families(text: str) -> list[str]:
    result: list[str] = []
    for label, patterns in CORE_PRODUCT_PATTERNS.items():
        if first_hit(text, patterns):
            result.append(label)
    return result


def has_any(text: str, patterns: Iterable[str]) -> bool:
    return any(pattern in text for pattern in patterns)


def detect_issues(text: str, stage: str, product_families: list[str]) -> list[dict[str, str | float]]:
    issues: list[dict[str, str | float]] = []
    family_count = len(product_families)
    service_hit = has_any(text, SERVICE_PATTERNS)

    for issue_id, issue_name, patterns, recommendation in ISSUES:
        if issue_id == "qualification_gap":
            qualification_tokens = [
                "размер",
                "тираж",
                "срок",
                "доставка",
                "монтаж",
                "адрес",
                "материал",
                "оплата",
                "бюджет",
                "лпр",
                "кто принимает решение",
            ]
            qual_score = sum(1 for token in qualification_tokens if token in text)
            if qual_score < 1:
                issues.append(
                    {
                        "issue_id": issue_id,
                        "issue_name": issue_name,
                        "trigger": "мало квалифицирующих вопросов",
                        "manager_action": "не добрался до обязательной квалификации",
                        "risk": stage_weight(stage),
                        "recommendation": recommendation,
                    }
                )
            continue

        if issue_id == "no_next_step":
            if stage_class(stage) != "positive" and not has_any(text, patterns):
                issues.append(
                    {
                        "issue_id": issue_id,
                        "issue_name": issue_name,
                        "trigger": "нет зафиксированного next step",
                        "manager_action": "не закрыл звонок на следующий шаг",
                        "risk": stage_weight(stage),
                        "recommendation": recommendation,
                    }
                )
            continue

        if issue_id == "upsell_missed":
            if family_count >= 1 and not service_hit:
                issues.append(
                    {
                        "issue_id": issue_id,
                        "issue_name": issue_name,
                        "trigger": "одиночная продажа без допуслуг",
                        "manager_action": "не предложил пакетирование и допродажу",
                        "risk": 0.5 if stage_class(stage) == "positive" else stage_weight(stage),
                        "recommendation": recommendation,
                    }
                )
            continue

        if has_any(text, patterns):
            trigger = first_hit(text, patterns)
            manager_action = {
                "price_objection": "не перевел разговор от цены к ценности",
                "postpone": "не дожал сделку до конкретной даты",
                "compare": "не отстроился от конкурентов",
            }[issue_id]
            risk = stage_weight(stage)
            if issue_id == "price_objection" and stage in {"Проиграли по цене", "Низкий бюджет"}:
                risk *= 1.4
            if issue_id == "postpone" and stage in {"Отложено", "Просрочено"}:
                risk *= 1.3
            if issue_id == "compare" and stage in {"Проиграли по цене", "Отказ ms"}:
                risk *= 1.2
            issues.append(
                {
                    "issue_id": issue_id,
                    "issue_name": issue_name,
                    "trigger": trigger,
                    "manager_action": manager_action,
                    "risk": risk,
                    "recommendation": recommendation,
                }
            )

    issues.sort(key=lambda item: float(item["risk"]), reverse=True)
    return issues


@dataclass
class CallRecord:
    row_number: int
    file_stem: str
    date_text: str
    month_key: str
    stage: str
    stage_group: str
    text: str
    issues: list[dict[str, str | float]]
    product_families: list[str]


def parse_month(date_text: str) -> str:
    try:
        dt = datetime.strptime(date_text, "%d.%m.%Y %H:%M:%S")
    except ValueError:
        return date_text[:7] if len(date_text) >= 7 else date_text
    return dt.strftime("%m.%Y")


def build_records(ws) -> list[CallRecord]:
    records: list[CallRecord] = []
    for row_idx in range(2, ws.max_row + 1):
        filename = ws[f"R{row_idx}"].value or ""
        file_stem = Path(str(filename)).name.replace(".mp3", "")
        json_path = INPUT_JSON_DIR / f"{file_stem}.json"
        text = read_json_text(json_path) if json_path.exists() else ""
        stage = str(ws[f"L{row_idx}"].value or "")
        product_families = extract_product_families((ws[f"K{row_idx}"].value or "") + " " + text)
        issues = detect_issues(text, stage, product_families)
        records.append(
            CallRecord(
                row_number=row_idx,
                file_stem=file_stem,
                date_text=str(ws[f"A{row_idx}"].value or ""),
                month_key=parse_month(str(ws[f"A{row_idx}"].value or "")),
                stage=stage,
                stage_group=stage_class(stage),
                text=text,
                issues=issues,
                product_families=product_families,
            )
        )
    return records


def enrich_calls_sheet(ws, records: list[CallRecord]) -> list[str]:
    new_headers = [
        "Ключевая проблема",
        "Возражение / триггер",
        "Что сделал / не сделал менеджер",
        "Риск",
        "Рекомендация",
        "Потенциал апсейла",
        "Категории сделки",
    ]
    start_col = ws.max_column + 1
    for offset, header in enumerate(new_headers, start=start_col):
        ws.cell(row=1, column=offset, value=header)

    for record in records:
        primary = record.issues[0] if record.issues else None
        issue_names = "; ".join(issue["issue_name"] for issue in record.issues[:3]) or "Нет критичных проблем"
        trigger = primary["trigger"] if primary else "Нет явного триггера"
        action = primary["manager_action"] if primary else "Сделка движется без выраженного узкого места"
        risk = round(float(primary["risk"]), 2) if primary else 0.0
        recommendation = primary["recommendation"] if primary else ""
        upsell = "Да" if "Упущенный апсейл" in issue_names else "Нет"
        categories = ", ".join(record.product_families) if record.product_families else "не распознано"
        values = [issue_names, trigger, action, risk, recommendation, upsell, categories]
        for idx, value in enumerate(values, start=start_col):
            ws.cell(row=record.row_number, column=idx, value=value)

    return new_headers


def style_calls_sheet(ws, total_cols: int) -> None:
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{ws.cell(row=1, column=total_cols).coordinate}"
    widths = {
        "A": 18,
        "B": 20,
        "C": 12,
        "D": 14,
        "E": 16,
        "F": 14,
        "G": 14,
        "H": 38,
        "I": 12,
        "J": 12,
        "K": 28,
        "L": 24,
        "M": 14,
        "N": 26,
        "O": 34,
        "P": 34,
        "Q": 16,
        "R": 42,
    }
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width
    for col_idx in range(19, total_cols + 1):
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = 24
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = PatternFill("solid", fgColor="D9E2F3")


def aggregate(records: list[CallRecord]):
    issue_stats = defaultdict(lambda: Counter())
    monthly_stats = defaultdict(lambda: Counter())

    for record in records:
        month = record.month_key
        stage_group = record.stage_group
        monthly_stats[month]["calls"] += 1
        monthly_stats[month][f"{stage_group}_calls"] += 1

        seen = set()
        for issue in record.issues:
            issue_id = str(issue["issue_id"])
            if issue_id in seen:
                continue
            seen.add(issue_id)
            issue_stats[issue_id]["frequency"] += 1
            issue_stats[issue_id][f"{stage_group}_calls"] += 1
            issue_stats[issue_id]["risk_score"] += float(issue["risk"])
            monthly_stats[month][issue_id] += 1

    return issue_stats, monthly_stats


def issue_display_name(issue_id: str) -> str:
    for stored_id, issue_name, _, _ in ISSUES:
        if stored_id == issue_id:
            return issue_name
    return issue_id


def issue_recommendation(issue_id: str) -> str:
    for stored_id, _, _, recommendation in ISSUES:
        if stored_id == issue_id:
            return recommendation
    return ""


def build_summary_sheet(ws, records: list[CallRecord], issue_stats, monthly_stats) -> None:
    ws.delete_rows(1, ws.max_row)
    ws["A1"] = "Параметр"
    ws["B1"] = "Значение"
    meta = [
        ("Период", "01.03.2026 — 05.06.2026"),
        ("Мин. длительность", ">60 сек"),
        ("Только с записью", "Да"),
        ("Всего звонков", len(records)),
        ("Позитивные звонки", sum(1 for r in records if r.stage_group == "positive")),
        ("Негативные звонки", sum(1 for r in records if r.stage_group == "negative")),
        ("Доля позитивных звонков", f"{sum(1 for r in records if r.stage_group == 'positive') / len(records):.1%}"),
        ("Доля комплексных сделок", f"{sum(1 for r in records if len(r.product_families) >= 2) / len(records):.1%}"),
    ]
    row = 2
    for key, value in meta:
        ws.cell(row=row, column=1, value=key)
        ws.cell(row=row, column=2, value=value)
        row += 1

    start = row + 2
    headers = [
        "Ранг",
        "Проблема",
        "Частота",
        "Доля",
        "Вес риска",
        "Решение",
        "Позитив",
        "Нейтрал",
        "Негатив",
    ]
    for col, header in enumerate(headers, start=1):
        ws.cell(row=start, column=col, value=header)

    ranked = sorted(issue_stats.items(), key=lambda item: (item[1]["risk_score"], item[1]["frequency"]), reverse=True)
    summary_rows = []
    total_calls = len(records)
    for rank, (issue_id, stats) in enumerate(ranked, start=1):
        freq = stats["frequency"]
        score = round(stats["risk_score"], 1)
        summary_rows.append(
            [
                rank,
                issue_display_name(issue_id),
                freq,
                f"{freq / total_calls:.1%}",
                score,
                issue_recommendation(issue_id),
                stats["positive_calls"],
                stats["neutral_calls"],
                stats["negative_calls"],
            ]
        )

    for r_offset, values in enumerate(summary_rows, start=start + 1):
        for c_offset, value in enumerate(values, start=1):
            ws.cell(row=r_offset, column=c_offset, value=value)

    for cell in ws[start]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9EAD3")

    chart1 = BarChart()
    chart1.type = "bar"
    chart1.style = 10
    chart1.title = "Ключевые проблемы"
    chart1.y_axis.title = "Проблема"
    chart1.x_axis.title = "Вес риска"
    data = Reference(ws, min_col=5, min_row=start, max_row=start + min(8, len(summary_rows)))
    cats = Reference(ws, min_col=2, min_row=start + 1, max_row=start + min(8, len(summary_rows)))
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.height = 8
    chart1.width = 14
    ws.add_chart(chart1, "K2")

    trend_start = start + len(summary_rows) + 3
    ws.cell(row=trend_start, column=1, value="Месяц")
    ws.cell(row=trend_start, column=2, value="Проблемные звонки")
    ws.cell(row=trend_start, column=3, value="Скользящее среднее")

    month_order = sorted(monthly_stats.keys(), key=lambda m: datetime.strptime("01." + m, "%d.%m.%Y"))
    monthly_values = []
    for month in month_order:
        total_problem_calls = monthly_stats[month]["positive_calls"] + monthly_stats[month]["neutral_calls"] + monthly_stats[month]["negative_calls"]
        # here all calls are counted; problems are the calls with at least one detected issue
        problem_calls = sum(
            1
            for record in records
            if record.month_key == month and record.issues
        )
        monthly_values.append(problem_calls)

    for idx, month in enumerate(month_order, start=trend_start + 1):
        ws.cell(row=idx, column=1, value=month)
        value = sum(1 for record in records if record.month_key == month and record.issues)
        ws.cell(row=idx, column=2, value=value)
        if idx - (trend_start + 1) < 2:
            ws.cell(row=idx, column=3, value=value)
        else:
            prev2 = ws.cell(row=idx - 1, column=2).value or 0
            prev1 = ws.cell(row=idx - 2, column=2).value or 0
            ws.cell(row=idx, column=3, value=round((prev2 + prev1 + value) / 3, 2))

    chart2 = LineChart()
    chart2.title = "Тренд проблемных звонков"
    chart2.y_axis.title = "Кол-во звонков"
    chart2.x_axis.title = "Месяц"
    data = Reference(ws, min_col=2, min_row=trend_start, max_row=trend_start + len(month_order))
    cats = Reference(ws, min_col=1, min_row=trend_start + 1, max_row=trend_start + len(month_order))
    chart2.add_data(data, titles_from_data=True)
    chart2.set_categories(cats)
    chart2.height = 7
    chart2.width = 14
    if chart2.series:
        chart2.series[0].trendline = Trendline(trendlineType="linear")
    ws.add_chart(chart2, "K20")

    chart3 = LineChart()
    chart3.title = "Скользящее среднее"
    chart3.y_axis.title = "Кол-во звонков"
    chart3.x_axis.title = "Месяц"
    data = Reference(ws, min_col=3, min_row=trend_start, max_row=trend_start + len(month_order))
    chart3.add_data(data, titles_from_data=True)
    chart3.set_categories(cats)
    chart3.height = 7
    chart3.width = 14
    ws.add_chart(chart3, "K35")

    for col in range(1, 10):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = 22
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["F"].width = 60
    ws.freeze_panes = "A2"
    for row_cells in ws.iter_rows():
        for cell in row_cells:
            if cell.value is not None and cell.row != 1:
                cell.alignment = Alignment(vertical="top", wrap_text=True)


def build_forecast_sheet(ws, records: list[CallRecord], issue_stats) -> None:
    ws.delete_rows(1, ws.max_row)
    ws["A1"] = "Метрика"
    ws["B1"] = "Текущее"
    ws["C1"] = "Консервативно"
    ws["D1"] = "Базово"
    ws["E1"] = "Агрессивно"

    total = len(records)
    positive = sum(1 for r in records if r.stage_group == "positive")
    complex_calls = sum(1 for r in records if len(r.product_families) >= 2)
    single_core_calls = sum(1 for r in records if len(r.product_families) == 1)

    current_conv = positive / total
    current_complex = complex_calls / total
    current_single = single_core_calls / total

    price_share = issue_stats["price_objection"]["frequency"] / total if issue_stats["price_objection"]["frequency"] else 0
    postpone_share = issue_stats["postpone"]["frequency"] / total if issue_stats["postpone"]["frequency"] else 0
    qual_share = issue_stats["qualification_gap"]["frequency"] / total if issue_stats["qualification_gap"]["frequency"] else 0
    upsell_share = issue_stats["upsell_missed"]["frequency"] / total if issue_stats["upsell_missed"]["frequency"] else 0

    rows = [
        ("Конверсия звонок→сделка", current_conv, current_conv + min(0.035, price_share * 0.20 * 0.25 + postpone_share * 0.15 * 0.20), current_conv + min(0.060, price_share * 0.30 * 0.35 + postpone_share * 0.20 * 0.25 + qual_share * 0.15 * 0.15), current_conv + min(0.100, price_share * 0.40 * 0.45 + postpone_share * 0.25 * 0.30 + qual_share * 0.20 * 0.25)),
        ("Доля комплексных сделок", current_complex, current_complex + min(0.050, current_single * 0.25), current_complex + min(0.080, current_single * 0.35), current_complex + min(0.120, current_single * 0.45)),
        ("Доля звонков с апсейлом", upsell_share, upsell_share + 0.03, upsell_share + 0.05, upsell_share + 0.08),
    ]

    for row_idx, (metric, current, conservative, base, aggressive) in enumerate(rows, start=2):
        ws.cell(row=row_idx, column=1, value=metric)
        ws.cell(row=row_idx, column=2, value=round(current, 1) if current > 1 else f"{current:.1%}")
        ws.cell(row=row_idx, column=3, value=f"{conservative:.1%}")
        ws.cell(row=row_idx, column=4, value=f"{base:.1%}")
        ws.cell(row=row_idx, column=5, value=f"{aggressive:.1%}")

    assumptions_row = 7
    ws.cell(row=assumptions_row, column=1, value="Допущения")
    ws.cell(row=assumptions_row + 1, column=1, value="Снижение price objection")
    ws.cell(row=assumptions_row + 1, column=2, value="20% / 30% / 40%")
    ws.cell(row=assumptions_row + 2, column=1, value="Снижение postponed deals")
    ws.cell(row=assumptions_row + 2, column=2, value="15% / 20% / 25%")
    ws.cell(row=assumptions_row + 3, column=1, value="Рост complex-sale share")
    ws.cell(row=assumptions_row + 3, column=2, value="25% / 35% / 45% от single-core сделок")

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="FCE5CD")
    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 16


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    wb = load_workbook(INPUT_XLSX)
    ws_calls = wb["Звонки"]
    ws_summary = wb["Сводка"]

    records = build_records(ws_calls)
    new_headers = enrich_calls_sheet(ws_calls, records)
    style_calls_sheet(ws_calls, ws_calls.max_column)

    issue_stats, monthly_stats = aggregate(records)

    build_summary_sheet(ws_summary, records, issue_stats, monthly_stats)
    if "Прогноз" in wb.sheetnames:
        del wb["Прогноз"]
    ws_forecast = wb.create_sheet("Прогноз")
    build_forecast_sheet(ws_forecast, records, issue_stats)

    wb.save(OUTPUT_XLSX)
    print(f"saved {OUTPUT_XLSX}")


if __name__ == "__main__":
    main()
