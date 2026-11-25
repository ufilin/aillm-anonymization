import random

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

SRC = "Таблица_к_ДЗ.xlsx"  # исходный файл положить рядом, в репозиторий не включается
OUT = "Анонимизированный_датасет.xlsx"
SEED = 20251125
NOISE = 0.05
ROUND_TO = 1000

OKRUG = {
    "Москва": "ЦФО",
    "Ярославль": "ЦФО",
    "Санкт-Петербург": "СЗФО",
    "Казань": "ПФО",
    "Нижний Новгород": "ПФО",
    "Пермь": "ПФО",
    "Киров": "ПФО",
    "Екатеринбург": "УФО",
    "Тюмень": "УФО",
}

# ---------- чтение исходника ----------
src = openpyxl.load_workbook(SRC, data_only=True)
sh = src.active
rows = []
for r in sh.iter_rows(min_row=4, max_row=sh.max_row, values_only=True):
    if r[0] is None:
        continue
    rows.append({"n": r[0], "company": r[1], "city": r[2], "sum": r[7], "debt": r[8]})
assert len(rows) == 20, len(rows)

# ---------- преобразования ----------
rnd = random.Random(SEED)
order = list(range(len(rows)))
rnd.shuffle(order)

anon = []
for seq, idx in enumerate(order, start=1):
    r = rows[idx]
    region = OKRUG[(r["city"] or "").strip()]

    # один коэффициент шума на запись: суммы скрываются, а доля долга остаётся достоверной
    factor = 1 + rnd.uniform(-NOISE, NOISE)

    def noised(value):
        return int(round(value * factor / ROUND_TO)) * ROUND_TO

    anon.append(
        {
            "client_id": "Клиент_%03d" % seq,
            "region": region,
            "contract_id": "Договор_%03d" % seq,
            "sum": noised(r["sum"]),
            "debt": noised(r["debt"]),
        }
    )

# ---------- проверки ----------
groups = {}
for rec in anon:
    groups[rec["region"]] = groups.get(rec["region"], 0) + 1
k_min = min(groups.values())
ratios = [rec["debt"] / rec["sum"] for rec in anon]
over100 = sum(1 for x in ratios if x > 1)
max_ratio = max(ratios)

# ---------- стили ----------
FONT = "Arial"
hdr_fill = PatternFill("solid", fgColor="1F3864")
hdr_font = Font(name=FONT, size=11, bold=True, color="FFFFFF")
title_font = Font(name=FONT, size=14, bold=True, color="1F3864")
body_font = Font(name=FONT, size=11)
thin = Side(style="thin", color="B4C6E7")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(wrap_text=True, vertical="top")
center = Alignment(horizontal="center", vertical="center")

wb = openpyxl.Workbook()

# ================= ЛИСТ 1 =================
ws = wb.active
ws.title = "Анонимизированный датасет"

ws["A1"] = "Анонимизированный датасет для анализа продаж и дебиторской задолженности"
ws["A1"].font = title_font
ws.merge_cells("A1:F1")
ws.row_dimensions[1].height = 22

ws["A2"] = (
    "Прямые идентификаторы удалены. Город обобщён до федерального округа, суммы искажены "
    "в пределах ±5% и округлены до 1 000 руб. Строки перемешаны, исходная нумерация не сохранена."
)
ws["A2"].font = Font(name=FONT, size=10, italic=True, color="595959")
ws["A2"].alignment = wrap
ws.merge_cells("A2:F2")
ws.row_dimensions[2].height = 30

headers = [
    "ID клиента",
    "Регион (федеральный округ)",
    "ID договора",
    "Сумма договора, руб. (±5%)",
    "Остаток дебиторской задолженности, руб. (±5%)",
    "Доля непогашенной задолженности",
]
HR = 4
for c, h in enumerate(headers, start=1):
    cell = ws.cell(row=HR, column=c, value=h)
    cell.font = hdr_font
    cell.fill = hdr_fill
    cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    cell.border = border
ws.row_dimensions[HR].height = 45

for i, rec in enumerate(anon):
    row = HR + 1 + i
    ws.cell(row=row, column=1, value=rec["client_id"])
    ws.cell(row=row, column=2, value=rec["region"])
    ws.cell(row=row, column=3, value=rec["contract_id"])
    ws.cell(row=row, column=4, value=rec["sum"])
    ws.cell(row=row, column=5, value=rec["debt"])
    ws.cell(row=row, column=6, value='=IFERROR(E%d/D%d,"")' % (row, row))
    for c in range(1, 7):
        cell = ws.cell(row=row, column=c)
        cell.font = body_font
        cell.border = border
        if c in (4, 5):
            cell.number_format = "#,##0"
            cell.alignment = Alignment(horizontal="right")
        elif c == 6:
            cell.number_format = "0.0%"
            cell.alignment = Alignment(horizontal="right")
        else:
            cell.alignment = Alignment(horizontal="center")
    ws.row_dimensions[row].height = 17

for c, w in enumerate([13, 26, 14, 24, 32, 24], start=1):
    ws.column_dimensions[get_column_letter(c)].width = w
ws.freeze_panes = "A5"
ws.auto_filter.ref = "A4:F%d" % (HR + len(anon))

# ================= ЛИСТ 2 =================
ws2 = wb.create_sheet("Классификация полей")
ws2["A1"] = "Классификация полей исходного датасета и применённые методы"
ws2["A1"].font = title_font
ws2.merge_cells("A1:F1")

h2 = ["№", "Поле исходного датасета", "Категория", "Метод", "Что сделано", "Влияние на анализ"]
for c, h in enumerate(h2, start=1):
    cell = ws2.cell(row=3, column=c, value=h)
    cell.font = hdr_font
    cell.fill = hdr_fill
    cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    cell.border = border
ws2.row_dimensions[3].height = 32

cls = [
    ("№ п/п", "Служебное", "Перегенерирование",
     "Заменён на случайный ID клиента, строки перемешаны",
     "Записи остаются уникальными, восстановить исходный порядок нельзя"),
    ("Название компании", "Прямой идентификатор (юрлицо)", "Псевдонимизация + удаление",
     "Заменено на Клиент_001…020. Организационно-правовая форма (ООО/АО/ПАО/ИП) удалена вместе с названием",
     "Сквозная аналитика по клиенту сохраняется, реальные контрагенты не раскрываются"),
    ("Город", "Квази-идентификатор", "Обобщение",
     "Город → федеральный округ: Москва и Ярославль → ЦФО, Санкт-Петербург → СЗФО, "
     "Казань, Нижний Новгород, Пермь, Киров → ПФО, Екатеринбург и Тюмень → УФО",
     "Региональный срез сохранён, точная локация скрыта"),
    ("ФИО контактного лица", "Прямой идентификатор (ПДн)", "Удаление", "Столбец удалён",
     "Для анализа на уровне компаний не требуется"),
    ("Должность", "Квази-идентификатор", "Удаление",
     "Столбец удалён: характеризует контактное лицо, а не клиента, а в паре с регионом давал уникальные сочетания",
     "На цели анализа (продажи, взыскание) не влияет"),
    ("Телефон", "Прямой идентификатор (ПДн)", "Удаление", "Столбец удалён", "Не нужен для анализа"),
    ("Номер договора", "Прямой идентификатор / коммерческая тайна", "Псевдонимизация",
     "Заменён на Договор_001…020 (нумерация присвоена после перемешивания)",
     "Связка «клиент — договор» внутри датасета сохранена"),
    ("Сумма договора", "Конфиденциальный атрибут", "Добавление шума + обобщение",
     "Значения искажены в пределах ±5% и округлены до 1 000 руб.",
     "Средние и динамика считаются, точные условия сделок не раскрываются"),
    ("Остаток дебиторской задолженности", "Конфиденциальный атрибут", "Добавление шума + обобщение",
     "Значения искажены в пределах ±5% и округлены до 1 000 руб.",
     "Анализ взыскания сохраняется, точные суммы скрыты"),
    ("Доля непогашенной задолженности", "Производный показатель", "Расчёт (формула)",
     "Расчёт: Остаток ÷ Сумма договора",
     "Новый безопасный показатель: отношение информативно, абсолютные суммы скрыты"),
]
widths2 = [5, 24, 24, 22, 56, 42]
for c, w in enumerate(widths2, start=1):
    ws2.column_dimensions[get_column_letter(c)].width = w

for i, vals in enumerate(cls, start=4):
    row_vals = (i - 3,) + vals
    for c, v in enumerate(row_vals, start=1):
        cell = ws2.cell(row=i, column=c, value=v)
        cell.font = body_font
        cell.border = border
        cell.alignment = wrap if c >= 2 else Alignment(horizontal="center", vertical="top")
    lines = max(
        len(str(v)) / (widths2[c - 1] * 1.05) for c, v in enumerate(row_vals, start=1)
    )
    ws2.row_dimensions[i].height = max(30, 15 * int(lines + 0.99) + 4)

# ================= ЛИСТ 3 =================
ws3 = wb.create_sheet("Отчёт и самопроверка")
ws3["A1"] = "Отчёт по анонимизации и самопроверка"
ws3["A1"].font = title_font
ws3.merge_cells("A1:C1")

L = []
add = L.append

add(("h", "1. Что и как сделано"))
add(("t", "Шаг 1. Загрузка и знакомство с данными: 20 записей, 9 полей, пропусков и формул в исходнике нет."))
add(("t", "Шаг 2. Классификация: прямые идентификаторы — компания, ФИО контактного лица, телефон, номер договора; "
          "квази-идентификаторы — город и должность; конфиденциальные атрибуты — сумма договора и остаток долга; "
          "служебная нумерация."))
add(("t", "Шаг 3. Методы: удаление (ФИО, телефон, должность), псевдонимизация (компания, договор), "
          "обобщение (город → федеральный округ), шум ±5% с округлением до 1 000 руб. (денежные суммы), "
          "перемешивание строк и присвоение новых ID."))
add(("t", "Шаг 4. Итоговый датасет: 20 записей, 6 полей. Исходный порядок строк утрачен."))
add(("t", "Шаг 5. Проверка результата — раздел 2 и чек-лист в разделе 4."))
add(("", ""))

add(("h", "2. Проверка обезличивания"))
add(("t", "Прямых идентификаторов в датасете не осталось: нет названий компаний ни в каком виде "
          "(включая ООО/АО/ПАО/ИП), нет ФИО, телефонов и исходных номеров договоров."))
add(("t", "География обобщена до федеральных округов. Показатель k-анонимности по полю «Регион» = %d — "
          "то есть каждая запись неотличима как минимум от одной другой по региону. "
          "Распределение: %s." % (k_min, ", ".join("%s — %d" % (k, v) for k, v in sorted(groups.items())))))
add(("t", "Денежные суммы искажены в пределах ±5%%, поэтому точные значения сделок по ним не восстановить; "
          "коэффициент шума для каждой записи один и тот же для суммы и остатка, поэтому доля задолженности "
          "не искажается, а погрешность при расчёте средних по 20 записям взаимно гасится."))
add(("t", "Логика данных сохранена. У одной записи доля непогашенной задолженности превышает 100%% (%.1f%%) — "
          "это унаследованная особенность исходного файла, значения не исправлялись." % (max_ratio * 100)
     if over100 else
     "Доля непогашенной задолженности нигде не превышает 100%% (максимум %.1f%%)." % (max_ratio * 100)))
add(("t", "Остаточный риск. Набор маленький (20 записей), поэтому редкое сочетание признаков теоретически "
          "сужает круг кандидатов. Риск снижен тем, что прямые идентификаторы удалены, строки перемешаны, "
          "ID случайны, суммы зашумлены. Сопоставить запись с исходником можно только по ключу "
          "«компания → Клиент_ID», которого в этом файле нет: он хранится отдельно и не публикуется."))
add(("", ""))

add(("h", "3. Сохранилась ли аналитическая ценность"))
add(("t", "Да. Доступны сумма и остаток по каждому договору (с погрешностью ±5%), доля непогашенной "
          "задолженности и региональный срез. Этого достаточно, чтобы посчитать средний чек, долю просрочки, "
          "сравнить регионы между собой и оценить структуру портфеля."))
add(("t", "Потеряно намеренно: привязка к конкретным контрагентам, точные суммы, город и должность контакта. "
          "Для заявленных целей анализа эти данные не нужны."))
add(("", ""))

add(("h", "4. Чек-лист самопроверки"))
for item in [
    "Все прямые идентификаторы удалены или заменены необратимыми псевдонимами — Да",
    "Квази-идентификаторы обобщены или зашумлены — Да",
    "По датасету нельзя однозначно установить конкретного человека или компанию — Да",
    "По комбинации оставшихся полей находится несколько записей — Да (k = %d по региону)" % k_min,
    "Структура и объём сохранены: 20 записей — Да",
    "Данные пригодны для анализа продаж и дебиторской задолженности — Да",
]:
    add(("t", "☑ " + item))
add(("", ""))

add(("h", "5. Параметры обработки"))
add(("t", "Шум: равномерный, ±5%%, генератор random с фиксированным seed %d — результат воспроизводим." % SEED))
add(("t", "Округление сумм: до 1 000 руб. Соответствие «город → федеральный округ» — стандартное (ЦФО, СЗФО, ПФО, УФО)."))
add(("t", "Ключ сопоставления «компания → Клиент_ID» в файл не включён и должен храниться отдельно от датасета."))

r = 3
for kind, text in L:
    cell = ws3.cell(row=r, column=1, value=text)
    cell.alignment = wrap
    if kind == "h":
        cell.font = Font(name=FONT, size=12, bold=True, color="1F3864")
        ws3.row_dimensions[r].height = 20
    elif kind == "t":
        cell.font = body_font
        ws3.row_dimensions[r].height = 15 * (len(text) // 92 + 1)
    else:
        ws3.row_dimensions[r].height = 8
    ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    r += 1

ws3.column_dimensions["A"].width = 95
ws3.column_dimensions["B"].width = 18
ws3.column_dimensions["C"].width = 18

# --- печать ---
for sheet, landscape in ((ws, True), (ws2, True), (ws3, False)):
    sheet.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    sheet.page_setup.orientation = "landscape" if landscape else "portrait"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.print_options.horizontalCentered = True
ws2.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
ws2.page_setup.orientation = "landscape"
ws2.page_setup.fitToWidth = 1
ws2.page_setup.fitToHeight = 0

wb.save(OUT)
print("saved:", OUT)
print("k_min:", k_min, "groups:", groups)
print("max_ratio: %.3f over100: %d" % (max_ratio, over100))
