"""Збирає BEEF_Графік_Жовтень_2026.xlsx із schedule.json (див. generate.py)."""
import calendar
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

HERE = Path(__file__).parent
data = json.loads((HERE / "schedule.json").read_text(encoding="utf-8"))
YEAR, MONTH = data["year"], data["month"]
DAYS = calendar.monthrange(YEAR, MONTH)[1]
WDN = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Нд"]
WD = [calendar.weekday(YEAR, MONTH, d) for d in range(1, DAYS + 1)]
CODES = ["Р", "С", "В", "П"]
CODE_NAMES = {"Р": "Ранкова", "С": "Середня (денний пік)", "В": "Вечірня", "П": "Пікова вечірня"}
TIMES = {
    "Р": ["09:30–17:00"] * 7,
    "С": ["12:00–20:00"] * 7,
    "В": ["15:00–22:30", "15:00–22:30", "16:00–23:30", "16:00–23:30",
          "16:00–23:30", "16:00–23:30", "15:30–22:30"],
    "П": ["—", "—", "—", "17:30–23:30", "17:30–23:30", "17:30–23:30", "—"],
}
PURPOSE = {
    "Р": "Відкриття, підготовка залу/цехів, ланч, бізнес-ланч",
    "С": "Підсилення денного піку та раннього вечора (Чт–Нд)",
    "В": "Вечірній сервіс, закриття зміни, прибирання, касові звіти",
    "П": "Підсилення вечірнього піку 19:00–22:00 і закриття (Чт–Сб)",
}
LOAD = [  # рівень, пояснення — ПРИПУЩЕННЯ до отримання даних POS
    ("Низький", "Повільний старт тижня; основний потік — ділові ланчі"),
    ("Низький", "Аналогічно понеділку"),
    ("Середній", "Перший день подовженого режиму до 23:00; вечір помірний"),
    ("Високий", "Ділові вечері, початок «вечорів вихідних»"),
    ("Піковий", "Максимальна вечірня посадка, банкети, винні замовлення"),
    ("Піковий", "Високий день і найвищий вечір тижня"),
    ("Високий", "Сімейний ланч і ранній вечір; закриття о 22:00"),
]

FONT = "Arial"
f_base = Font(name=FONT, size=10)
f_bold = Font(name=FONT, size=10, bold=True)
f_title = Font(name=FONT, size=14, bold=True)
f_sub = Font(name=FONT, size=10, italic=True, color="595959")
f_input = Font(name=FONT, size=10, color="0000FF")
f_head = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_small = Font(name=FONT, size=8, color="808080")
fill_head = PatternFill("solid", fgColor="3A2A20")
fill_sect = PatternFill("solid", fgColor="E8E1D9")
fill_wkend = PatternFill("solid", fgColor="F3EEE8")
fill_input = PatternFill("solid", fgColor="FFF9C4")
FILL_CODE = {"Р": "FCE4B6", "С": "D5E8D4", "В": "CFE0F3", "П": "E4D4EE"}
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center")
wrap = Alignment(wrap_text=True, vertical="top")

wb = Workbook()


def put(ws, ref, value, font=f_base, fill=None, align=None, fmt=None, b=False):
    c = ws[ref]
    c.value = value
    c.font = font
    if fill:
        c.fill = fill
    if align:
        c.alignment = align
    if fmt:
        c.number_format = fmt
    if b:
        c.border = border
    return c


# ---------------------------------------------------------------- Інструкція
ws = wb.active
ws.title = "Інструкція"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 110
put(ws, "B2", "BEEF — графік роботи персоналу, жовтень 2026", f_title)
put(ws, "B3", "Робочий документ для менеджменту. Версія 1.0. Чисельність і норми — проєктні, підлягають "
              "калібруванню за фактичними даними.", f_sub)
lines = [
    ("Режим роботи ресторану", True),
    ("Неділя, понеділок, вівторок — 10:00–22:00. Середа, четвер, п'ятниця, субота — 10:00–23:00.", False),
    ("Як побудовано графік", True),
    ("1. Завантаження оцінено за днями тижня (аркуш «Норми»): Пн–Вт низьке, Ср середнє, Чт і Нд високе, Пт–Сб пікове.", False),
    ("2. Для кожної посади задано норму: скільки людей потрібно на кожну зміну в кожен день тижня.", False),
    ("3. Замість двох довгих змін використано чотири типи змін (аркуш «Зміни»): ранкова, середня, вечірня та пікова. "
     "Це дає змогу додавати людей лише в години піку, а не на весь день — фонд годин витрачається там, де є гості.", False),
    ("4. Розподіл змін між співробітниками виконано алгоритмом оптимізації з жорсткими правилами (нижче). "
     "Години в межах посади вирівняні, п'ятниці та суботи розподілені рівномірно.", False),
    ("Обов'язкові правила графіка", True),
    ("• Після вечірньої (В), середньої (С) або пікової (П) зміни наступного дня ранкова (Р) зміна заборонена. "
     "Причина: ст. 59 КЗпП — перерва між змінами не менше подвійної тривалості попередньої зміни; крім того, "
     "працівник після закриття о 23:30 не може якісно відкрити зал о 09:30.", False),
    ("• Не більше 5 робочих днів поспіль (щотижневий безперервний відпочинок, ст. 70 КЗпП).", False),
    ("• Вихідні — переважно парами; кожен співробітник має щонайменше один вихідний у Пт/Сб/Нд за місяць.", False),
    ("Як користуватися файлом", True),
    ("• Жовті клітинки — для заповнення: замініть «Офіціант 1» тощо на прізвища та ім'я співробітників.", False),
    ("• Сині значення — вхідні параметри (норми, години змін, норма годин місяця). Їх можна змінювати.", False),
    ("• Коди змін у графіку можна змінювати вручну (Р, С, В, П; порожня клітинка — вихідний). "
     "Години, підсумки та рядки «Контроль» перераховуються автоматично.", False),
    ("• Рядок «Контроль» під кожною посадою: «✓» — покриття відповідає нормі; від'ємне число (червоне) — "
     "бракує людей на зміну; додатне (жовте) — людей більше за норму.", False),
    ("• Приклад: Офіціант 3 у суботу має «П» — він виходить о 17:30 і працює до 23:30, підсилюючи вечірній пік.", False),
    ("Що потрібно для точного калібрування (наступна версія)", True),
    ("1. Фактичні дані POS за останні 8–12 тижнів: кількість гостей (covers) і чеків по годинах і днях тижня.", False),
    ("2. Кількість посадкових місць і столів, зони обслуговування, плановані банкети жовтня.", False),
    ("3. Фактичний штат: прізвища, посади, ставки, відпустки, обмеження (навчання, сумісництво).", False),
    ("4. Цільовий стандарт навантаження: скільки столів/гостей закріплюється за одним офіціантом у пік.", False),
    ("Важливі примітки", True),
    ("• Чисельність команд у файлі — розрахункова (сума змін за місяць / ≈23,5 зміни на людину, але не менше "
     "за кількість людей у найнапруженіший день). Це не фактичний штат ресторану.", False),
    ("• Норма робочого часу жовтня 2026 при 40-годинному тижні — 176 год (22 робочі дні). Під час воєнного стану "
     "святкові дні не є вихідними. Значення та облік (рекомендовано підсумований облік робочого часу) "
     "необхідно підтвердити з бухгалтерією/HR.", False),
    ("• Години змін у файлі — тривалість зміни від початку до кінця. Порядок обліку перерви на харчування "
     "персоналу визначається внутрішніми правилами трудового розпорядку.", False),
    ("• Офіціанти, сомельє та бармени отримують на щоденному брифінгу (09:40 і 16:10) розподіл зон і стоп-лист; "
     "час брифінгу вже входить у зміну.", False),
]
r = 5
for text, is_head in lines:
    if is_head:
        r += 1
        put(ws, f"B{r}", text, Font(name=FONT, size=11, bold=True))
    else:
        put(ws, f"B{r}", text, f_base, align=Alignment(wrap_text=True, vertical="top"))
    r += 1

# ---------------------------------------------------------------- Зміни
wz = wb.create_sheet("Зміни")
wz.sheet_view.showGridLines = False
put(wz, "A1", "Типи змін, час і години", f_title)
put(wz, "A2", "Час роботи ресторану: Нд–Вт 10:00–22:00; Ср–Сб 10:00–23:00. Персонал приходить за 30 хв "
              "до відкриття і залишається 30 хв після закриття.", f_sub)
wz.column_dimensions["A"].width = 8
wz.column_dimensions["B"].width = 24
for i in range(7):
    wz.column_dimensions[L(3 + i)].width = 13
wz.column_dimensions["J"].width = 60
hdr = ["Код", "Зміна"] + WDN + ["Призначення"]
for i, h in enumerate(hdr):
    put(wz, f"{L(1 + i)}4", h, f_head, fill_head, center, b=True)
for k, code in enumerate(CODES):
    row = 5 + k
    put(wz, f"A{row}", code, f_bold, PatternFill("solid", fgColor=FILL_CODE[code]), center, b=True)
    put(wz, f"B{row}", CODE_NAMES[code], b=True)
    for i in range(7):
        put(wz, f"{L(3 + i)}{row}", TIMES[code][i], align=center, b=True)
    put(wz, f"J{row}", PURPOSE[code], b=True)

put(wz, "A11", "Оплачувані години зміни (вхідні дані)", f_bold)
for i, h in enumerate(["Код", "Зміна"] + WDN):
    put(wz, f"{L(1 + i)}12", h, f_head, fill_head, center, b=True)
HOURS_ROW = {}
for k, code in enumerate(CODES):
    row = 13 + k
    HOURS_ROW[code] = row
    put(wz, f"A{row}", code, f_bold, PatternFill("solid", fgColor=FILL_CODE[code]), center, b=True)
    put(wz, f"B{row}", CODE_NAMES[code], b=True)
    for i in range(7):
        put(wz, f"{L(3 + i)}{row}", data["hours"][code][i], f_input, align=center, fmt="0.0", b=True)

put(wz, "A19", "Норма робочого часу на місяць, год", f_bold)
c = put(wz, "C19", 176, f_input, fill_input, center, "0", b=True)
c.comment = Comment("22 робочі дні жовтня 2026 × 8 год (40-годинний тиждень). "
                    "Підтвердити з бухгалтерією/HR.", "BEEF")
NORM_REF = "Зміни!$C$19"

put(wz, "A21", "Правила поєднання змін", f_bold)
rules = [
    "Р → будь-яка зміна наступного дня: дозволено.",
    "С, В, П → Р наступного дня: заборонено (міжзмінний відпочинок, ст. 59 КЗпП). Після вечора — вечір, середня або вихідний.",
    "Не більше 5 робочих днів поспіль.",
    "Під час зміни мають бути 30 хв перекриття «ранок–вечір» для передачі зміни (16:00–17:00 / 15:00–17:00).",
]
for i, t in enumerate(rules):
    put(wz, f"A{22 + i}", "• " + t)

# ---------------------------------------------------------------- Норми
wn = wb.create_sheet("Норми")
wn.sheet_view.showGridLines = False
put(wn, "A1", "Профіль завантаження та норми укомплектування", f_title)
put(wn, "A2", "Рівні завантаження — експертне припущення до отримання даних POS. Норми (сині) можна змінювати; "
              "контроль покриття в графіках перерахується автоматично.", f_sub)
wn.column_dimensions["A"].width = 30
wn.column_dimensions["B"].width = 8
for i in range(7):
    wn.column_dimensions[L(3 + i)].width = 11
wn.column_dimensions["J"].width = 12
wn.column_dimensions["K"].width = 3
wn.column_dimensions["L"].width = 28  # прихований ключ — у кінці

put(wn, "A4", "День тижня", f_head, fill_head, center, b=True)
put(wn, "B4", "", f_head, fill_head, b=True)
for i in range(7):
    put(wn, f"{L(3 + i)}4", WDN[i], f_head, fill_head, center, b=True)
put(wn, "A5", "Рівень завантаження", f_bold, b=True)
put(wn, "A6", "Режим роботи", f_bold, b=True)
put(wn, "A7", "Днів у жовтні", f_bold, b=True)
for i in range(7):
    put(wn, f"{L(3 + i)}5", LOAD[i][0], f_input, align=center, b=True)
    put(wn, f"{L(3 + i)}6", "10–22" if i in (0, 1, 6) else "10–23", align=center, b=True)
    put(wn, f"{L(3 + i)}7", WD.count(i), align=center, b=True)
for row in (5, 6, 7):
    put(wn, f"B{row}", "", b=True)
put(wn, "A8", "Обґрунтування рівня", f_bold)
for i in range(7):
    put(wn, f"{L(3 + i)}8", LOAD[i][1], f_small, align=Alignment(wrap_text=True, vertical="top"))
wn.row_dimensions[8].height = 64

put(wn, "A10", "Норми: кількість людей на зміну", f_bold)
for i, h in enumerate(["Посада", "Зміна"] + WDN + ["Змін/міс."]):
    put(wn, f"{L(1 + i)}11", h, f_head, fill_head, center, b=True)
put(wn, "L11", "Ключ (службовий)", f_small)
NORM_FIRST = 12
r = NORM_FIRST
for dept, positions in data["demand"].items():
    put(wn, f"A{r}", dept, f_bold, fill_sect)
    for col in range(2, 11):
        wn.cell(r, col).fill = fill_sect
    r += 1
    for pos, norms in positions.items():
        for code in CODES:
            if code not in norms:
                continue
            put(wn, f"A{r}", pos, b=True)
            put(wn, f"B{r}", code, f_bold, PatternFill("solid", fgColor=FILL_CODE[code]), center, b=True)
            for i in range(7):
                put(wn, f"{L(3 + i)}{r}", norms[code][i], f_input, align=center, b=True)
            put(wn, f"J{r}", f"=SUMPRODUCT(C{r}:I{r},$C$7:$I$7)", f_bold, align=center, b=True)
            put(wn, f"L{r}", f'=A{r}&"|"&B{r}', f_small)
            r += 1
NORM_LAST = r - 1
NORM_TABLE = f"Норми!$C${NORM_FIRST}:$I${NORM_LAST}"
NORM_KEYS = f"Норми!$L${NORM_FIRST}:$L${NORM_LAST}"
NORM_DAYS = "Норми!$C$11:$I$11"
wn.freeze_panes = "C12"

# ---------------------------------------------------------------- Графіки
FIRST_D, LAST_D = 4, 3 + DAYS          # D..AH
C_SH, C_HR, C_OFF, C_WK, C_DEV = (LAST_D + i for i in range(1, 6))
DL, DR = L(FIRST_D), L(LAST_D)
summary_refs = {}  # (dept, pos) -> (sheet, first_row, last_row)
daily_rows = {}


def roster(dept, title):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 24
    for d in range(DAYS):
        ws.column_dimensions[L(FIRST_D + d)].width = 4.2
    for col, w in zip((C_SH, C_HR, C_OFF, C_WK, C_DEV), (7, 8, 9, 8, 10)):
        ws.column_dimensions[L(col)].width = w

    put(ws, "A1", f"BEEF — графік роботи: {dept.lower()}, жовтень 2026", f_title)
    put(ws, "A2", "Р 09:30–17:00 · С 12:00–20:00 · В Пн–Вт 15:00–22:30, Ср–Сб 16:00–23:30, Нд 15:30–22:30 · "
                  "П 17:30–23:30 (Чт–Сб) · порожньо — вихідний", f_sub)

    # шапка
    for col, h in ((1, "№"), (2, "Посада"), (3, "Співробітник")):
        put(ws, f"{L(col)}4", h, f_head, fill_head, center, b=True)
        put(ws, f"{L(col)}5", "", f_head, fill_head, b=True)
        put(ws, f"{L(col)}6", "", f_head, fill_head, b=True)
    put(ws, "C6", "Завантаження", f_head, fill_head, Alignment(horizontal="right"), b=True)
    for d in range(DAYS):
        col = L(FIRST_D + d)
        put(ws, f"{col}4", d + 1, f_head, fill_head, center, b=True)
        put(ws, f"{col}5", WDN[WD[d]], f_bold, fill_wkend if WD[d] in (4, 5) else None, center, b=True)
        put(ws, f"{col}6", f'=LEFT(INDEX(Норми!$C$5:$I$5,MATCH({col}$5,Норми!$C$4:$I$4,0)),3)',
            f_small, fill_wkend if WD[d] in (4, 5) else None, center, b=True)
    for col, h in zip((C_SH, C_HR, C_OFF, C_WK, C_DEV),
                      ("Змін", "Годин", "Вихідних", "Пт+Сб", "Δ до норми, год")):
        put(ws, f"{L(col)}4", h, f_head, fill_head, Alignment(horizontal="center", vertical="center",
                                                               wrap_text=True), b=True)
        ws.merge_cells(start_row=4, start_column=col, end_row=6, end_column=col)
    ws.row_dimensions[4].height = 18

    # службові рядки годин
    for k, code in enumerate(CODES):
        row = 7 + k
        put(ws, f"C{row}", f"год. зміни {code}", f_small, align=Alignment(horizontal="right"))
        for d in range(DAYS):
            col = L(FIRST_D + d)
            put(ws, f"{col}{row}", f"=INDEX(Зміни!$C${HOURS_ROW[code]}:$I${HOURS_ROW[code]},"
                                   f"MATCH({col}$5,Зміни!$C$12:$I$12,0))", f_small, align=center, fmt="0.#")
        ws.row_dimensions[row].outlineLevel = 1
        ws.row_dimensions[row].hidden = True

    r = 12
    first_body = r
    n = 0
    for pos, rows in data["departments"][dept].items():
        put(ws, f"A{r}", pos.upper(), f_bold, fill_sect)
        for col in range(2, C_DEV + 1):
            ws.cell(r, col).fill = fill_sect
        r += 1
        p_first = r
        for i, sched in enumerate(rows):
            n += 1
            put(ws, f"A{r}", n, align=center, b=True)
            put(ws, f"B{r}", pos, b=True)
            put(ws, f"C{r}", f"{pos} {i + 1}", f_input, fill_input, b=True)
            for d in range(DAYS):
                put(ws, f"{L(FIRST_D + d)}{r}", sched[d] or None, f_bold, align=center, b=True)
            rng = f"{DL}{r}:{DR}{r}"
            hrs = "+".join(f'SUMPRODUCT(({rng}="{c}")*${DL}${7 + k}:${DR}${7 + k})'
                           for k, c in enumerate(CODES))
            put(ws, f"{L(C_SH)}{r}", f"=COUNTA({rng})", align=center, b=True)
            put(ws, f"{L(C_HR)}{r}", "=" + hrs, f_bold, align=center, fmt="0.0", b=True)
            put(ws, f"{L(C_OFF)}{r}", f"=COLUMNS({rng})-{L(C_SH)}{r}", align=center, b=True)
            put(ws, f"{L(C_WK)}{r}", f'=SUMPRODUCT(({rng}<>"")*((${DL}$5:${DR}$5="Пт")+(${DL}$5:${DR}$5="Сб")))',
                align=center, b=True)
            put(ws, f"{L(C_DEV)}{r}", f"={L(C_HR)}{r}-{NORM_REF}", align=center, fmt='+0.0;-0.0;0', b=True)
            r += 1
        p_last = r - 1
        summary_refs[(dept, pos)] = (title, p_first, p_last)
        # контроль покриття
        for code in CODES:
            if code not in data["demand"][dept][pos]:
                continue
            put(ws, f"C{r}", f"Контроль {code} (факт − норма)", f_small, align=Alignment(horizontal="right"))
            for d in range(DAYS):
                col = L(FIRST_D + d)
                put(ws, f"{col}{r}",
                    f'=COUNTIF({col}${p_first}:{col}${p_last},"{code}")-INDEX({NORM_TABLE},'
                    f'MATCH("{pos}|{code}",{NORM_KEYS},0),MATCH({col}$5,{NORM_DAYS},0))',
                    f_small, align=center, fmt='+0;-0;"✓"')
            ws.conditional_formatting.add(
                f"{DL}{r}:{DR}{r}", CellIsRule(operator="lessThan", formula=["0"],
                                               fill=PatternFill("solid", fgColor="F4B6B6"),
                                               font=Font(name=FONT, size=8, bold=True, color="9C0006")))
            ws.conditional_formatting.add(
                f"{DL}{r}:{DR}{r}", CellIsRule(operator="greaterThan", formula=["0"],
                                               fill=PatternFill("solid", fgColor="FFEB9C")))
            r += 1
        r += 1
    last_body = r - 1

    # людино-години та кількість людей на день
    put(ws, f"C{r}", "Людей на зміні за день", f_bold, align=Alignment(horizontal="right"))
    for d in range(DAYS):
        col = L(FIRST_D + d)
        cond = "+".join(f'({col}{first_body}:{col}{last_body}="{c}")' for c in CODES)
        put(ws, f"{col}{r}", f"=SUMPRODUCT({cond})", f_bold, align=center, b=True)
    r += 1
    put(ws, f"C{r}", "Людино-години за день", f_bold, align=Alignment(horizontal="right"))
    for d in range(DAYS):
        col = L(FIRST_D + d)
        hrs = "+".join(f'SUMPRODUCT(({col}{first_body}:{col}{last_body}="{c}")*{col}${7 + k})'
                       for k, c in enumerate(CODES))
        put(ws, f"{col}{r}", "=" + hrs, f_bold, align=Alignment(horizontal="center", text_rotation=90),
            fmt="0", b=True)
    ws.row_dimensions[r].height = 30
    daily_rows[dept] = (title, r)
    put(ws, f"{L(C_SH)}{r}", "Разом:", f_bold, align=Alignment(horizontal="right"))
    put(ws, f"{L(C_HR)}{r}", f"=SUM({DL}{r}:{DR}{r})", f_bold, align=center, fmt="#,##0")

    # кольори кодів
    body = f"{DL}{first_body}:{DR}{last_body}"
    for code in CODES:
        ws.conditional_formatting.add(
            body, FormulaRule(formula=[f'{DL}{first_body}="{code}"'],
                              fill=PatternFill("solid", fgColor=FILL_CODE[code])))
    dev = f"{L(C_DEV)}{first_body}:{L(C_DEV)}{last_body}"
    ws.conditional_formatting.add(dev, CellIsRule(operator="greaterThan", formula=["0"],
                                                  font=Font(name=FONT, size=10, bold=True, color="C00000")))
    ws.freeze_panes = f"{DL}7"
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "4:6"


roster("ЗАЛ", "Графік Зал")
roster("КУХНЯ", "Графік Кухня")

# ---------------------------------------------------------------- Зведення
wsu = wb.create_sheet("Зведення")
wsu.sheet_view.showGridLines = False
put(wsu, "A1", "Зведення: чисельність і фонд робочого часу", f_title)
put(wsu, "A2", "Усі показники розраховуються з аркушів графіка.", f_sub)
widths = [14, 26, 10, 14, 12, 12, 12, 16]
for i, w in enumerate(widths):
    wsu.column_dimensions[L(1 + i)].width = w
hdr = ["Підрозділ", "Посада", "Людей", "Годин разом", "Середнє", "Мін.", "Макс.", "Сер. до норми"]
for i, h in enumerate(hdr):
    put(wsu, f"{L(1 + i)}4", h, f_head, fill_head, center, b=True)
r = 5
hc = L(C_HR)
for (dept, pos), (sheet, a, b_) in summary_refs.items():
    rng = f"'{sheet}'!{hc}{a}:{hc}{b_}"
    put(wsu, f"A{r}", dept, b=True)
    put(wsu, f"B{r}", pos, b=True)
    put(wsu, f"C{r}", f"=COUNTA('{sheet}'!C{a}:C{b_})", align=center, b=True)
    put(wsu, f"D{r}", f"=SUM({rng})", align=center, fmt="#,##0", b=True)
    put(wsu, f"E{r}", f"=IF(C{r}=0,0,D{r}/C{r})", align=center, fmt="0.0", b=True)
    put(wsu, f"F{r}", f"=MIN({rng})", align=center, fmt="0.0", b=True)
    put(wsu, f"G{r}", f"=MAX({rng})", align=center, fmt="0.0", b=True)
    put(wsu, f"H{r}", f"=E{r}-{NORM_REF}", align=center, fmt="+0.0;-0.0;0", b=True)
    r += 1
put(wsu, f"B{r}", "Разом", f_bold, b=True)
put(wsu, f"A{r}", "", b=True)
put(wsu, f"C{r}", f"=SUM(C5:C{r - 1})", f_bold, align=center, b=True)
put(wsu, f"D{r}", f"=SUM(D5:D{r - 1})", f_bold, align=center, fmt="#,##0", b=True)
total_row = r

r += 2
put(wsu, f"A{r}", "Фонд годин за днями тижня (середній день)", f_bold)
r += 1
for i, h in enumerate(["День", "Рівень", "Зал, год", "Кухня, год", "Разом, год"]):
    put(wsu, f"{L(1 + i)}{r}", h, f_head, fill_head, center, b=True)
r += 1
gs, gr = daily_rows["ЗАЛ"]
ks, kr = daily_rows["КУХНЯ"]
day_rng = f"${DL}$5:${DR}$5"
for i in range(7):
    put(wsu, f"A{r}", WDN[i], f_bold, align=center, b=True)
    put(wsu, f"B{r}", f"=INDEX(Норми!$C$5:$I$5,{i + 1})", align=center, b=True)
    for col, (s, rr) in (("C", (gs, gr)), ("D", (ks, kr))):
        put(wsu, f"{col}{r}",
            f"=SUMPRODUCT(('{s}'!{day_rng}=A{r})*'{s}'!${DL}${rr}:${DR}${rr})/COUNTIF('{s}'!{day_rng},A{r})",
            align=center, fmt="0", b=True)
    put(wsu, f"E{r}", f"=C{r}+D{r}", f_bold, align=center, fmt="0", b=True)
    r += 1

r += 1
put(wsu, f"A{r}", "Як читати: фонд годин має зростати разом із рівнем завантаження. Якщо після отримання даних POS "
                  "фактичні гості на годину персоналу суттєво різняться між днями — скоригуйте норми на аркуші «Норми».",
    f_sub)

wb.move_sheet("Зведення", offset=-2)
out = HERE / "BEEF_Графік_Жовтень_2026.xlsx"
wb.save(out)
print(out)
