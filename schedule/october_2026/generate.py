"""
Генератор графіка роботи BEEF на жовтень 2026.

Логіка:
1. Норми укомплектування (скільки людей на яку зміну в який день тижня)
   задаються в DEMAND і відображають профіль завантаження.
2. Початковий розподіл будується жадібно, далі покращується
   імітацією відпалу (обмін змінами між співробітниками однієї посади
   в один день — покриття при цьому не змінюється).
3. Жорсткі правила: міжзмінний відпочинок (ранкова зміна наступного дня
   можлива лише після ранкової зміни або вихідного), не більше 5 робочих
   днів поспіль.
4. Цілі оптимізації: рівномірні години в межах посади, вихідні парами,
   стабільність типу зміни, справедливий розподіл п'ятниць/субот.

Результат: schedule.json, який читає build_xlsx.py.
"""
import calendar
import json
import math
import random
from pathlib import Path

YEAR, MONTH = 2026, 10
DAYS = calendar.monthrange(YEAR, MONTH)[1]
WD = [calendar.weekday(YEAR, MONTH, d) for d in range(1, DAYS + 1)]  # 0 = Пн
CODES = ["Р", "С", "В", "П"]
MAX_CONSEC = 5
TARGET_SHIFTS = 23.5  # макс. середня к-ть змін на людину (≈176 год норми)

# Години оплачуваного часу зміни за днем тижня (Пн..Нд)
HOURS = {
    "Р": [7.5] * 7,                                # 09:30–17:00
    "С": [8.0] * 7,                                # 12:00–20:00
    "В": [7.5, 7.5, 7.5, 7.5, 7.5, 7.5, 7.0],      # Пн–Вт 15:00–22:30; Ср–Сб 16:00–23:30; Нд 15:30–22:30
    "П": [6.0] * 7,                                # 17:30–23:30 (лише Чт–Сб)
}

# Норми: посада -> код зміни -> кількість людей Пн..Нд
DEMAND = {
    "ЗАЛ": {
        "Менеджер зміни": {"Р": [1, 1, 1, 1, 1, 1, 1], "В": [1, 1, 1, 1, 1, 1, 1]},
        "Хостес":         {"Р": [1, 1, 1, 1, 1, 1, 1], "С": [0, 0, 0, 0, 0, 0, 1],
                           "В": [1, 1, 1, 1, 1, 1, 1], "П": [0, 0, 0, 0, 1, 1, 0]},
        "Офіціант":       {"Р": [2, 2, 2, 2, 3, 3, 3], "С": [0, 0, 0, 1, 1, 2, 2],
                           "В": [2, 2, 3, 3, 4, 4, 3], "П": [0, 0, 0, 1, 2, 2, 0]},
        "Ранер":          {"Р": [1, 1, 1, 1, 1, 2, 2], "С": [0, 0, 0, 0, 1, 1, 1],
                           "В": [1, 1, 1, 2, 2, 2, 1], "П": [0, 0, 0, 0, 1, 1, 0]},
        "Сомельє":        {"С": [0, 0, 0, 1, 1, 1, 0], "В": [1, 1, 1, 1, 1, 1, 1]},
        "Бармен":         {"Р": [1, 1, 1, 1, 1, 1, 1], "В": [1, 1, 1, 1, 1, 1, 1],
                           "П": [0, 0, 0, 0, 1, 1, 0]},
    },
    "КУХНЯ": {
        "Су-шеф":         {"Р": [1, 1, 1, 1, 1, 1, 1], "В": [1, 1, 1, 1, 1, 1, 1]},
        "Кухар гриль":    {"Р": [1, 1, 1, 1, 1, 1, 1], "С": [0, 0, 0, 0, 1, 1, 1],
                           "В": [1, 1, 1, 2, 2, 2, 1]},
        "Кухар гарячого цеху": {"Р": [1, 1, 1, 1, 1, 1, 1], "С": [0, 0, 0, 0, 0, 1, 1],
                           "В": [1, 1, 1, 1, 2, 2, 1]},
        "Кухар холодного цеху": {"Р": [1, 1, 1, 1, 1, 1, 1], "С": [0, 0, 0, 0, 1, 1, 0],
                           "В": [1, 1, 1, 1, 1, 1, 1]},
        "Кондитер":       {"Р": [1, 1, 1, 1, 1, 1, 1]},
        "Мийник":         {"Р": [1, 1, 1, 1, 1, 1, 1], "В": [1, 1, 1, 1, 1, 1, 1],
                           "П": [0, 0, 0, 0, 1, 1, 0]},
    },
}

LATE = {"С", "В", "П"}


def team_size(norms):
    total = sum(norms[c][WD[d]] for c in norms for d in range(DAYS))
    peak = max(sum(norms[c][WD[d]] for c in norms) for d in range(DAYS))
    return max(math.ceil(total / TARGET_SHIFTS), peak)


def hours(code, d):
    return HOURS[code][WD[d]] if code else 0.0


def violations(row):
    """Кількість порушень жорстких правил у рядку співробітника."""
    v = 0
    run = 0
    for d in range(DAYS):
        if row[d]:
            run += 1
            if run > MAX_CONSEC:
                v += 1
        else:
            run = 0
        if d and row[d] == "Р" and row[d - 1] in LATE:
            v += 1
    return v


def row_cost(row):
    c = 1000 * violations(row)
    # одиночні вихідні (вихідний, оточений робочими днями)
    for d in range(DAYS):
        if not row[d]:
            left = d == 0 or row[d - 1]
            right = d == DAYS - 1 or row[d + 1]
            if left and right and 0 < d < DAYS - 1:
                c += 6
    # зміна типу зміни між сусідніми робочими днями
    for d in range(1, DAYS):
        if row[d] and row[d - 1] and row[d] != row[d - 1]:
            c += 1.5
    # тиждень без вихідних-пари тощо не штрафуємо окремо
    return c


def group_cost(rows):
    c = sum(row_cost(r) for r in rows)
    hrs = [sum(hours(r[d], d) for d in range(DAYS)) for r in rows]
    mean = sum(hrs) / len(hrs)
    c += sum((h - mean) ** 2 for h in hrs) * 0.6
    # справедливість п'ятниць і субот (найприбутковіші й найважчі зміни)
    pk = [sum(1 for d in range(DAYS) if r[d] and WD[d] in (4, 5)) for r in rows]
    pm = sum(pk) / len(pk)
    c += sum((p - pm) ** 2 for p in pk) * 2
    # кожен має хоча б один вільний вихідний день (Пт/Сб/Нд) за місяць
    for r in rows:
        if all(r[d] for d in range(DAYS) if WD[d] in (4, 5, 6)):
            c += 30
    return c


def build(norms, n, rng):
    rows = [[""] * DAYS for _ in range(n)]
    for d in range(DAYS):
        need = []
        for code in CODES:
            need += [code] * norms.get(code, [0] * 7)[WD[d]]
        need.sort(key=lambda c: c != "Р")  # спершу ранкові — вони найобмеженіші
        free = set(range(n))
        for code in need:
            def score(e):
                r = rows[e]
                s = sum(1 for x in r[:d] if x) * 10
                run = 0
                for x in reversed(r[:d]):
                    if not x:
                        break
                    run += 1
                if run >= MAX_CONSEC:
                    s += 10000
                if code == "Р" and d and r[d - 1] in LATE:
                    s += 10000
                if d and r[d - 1] == code:
                    s -= 4
                return s + rng.random()
            e = min(free, key=score)
            rows[e][d] = code
            free.discard(e)
    return rows


def anneal(rows, rng, iters=60000):
    n = len(rows)
    cost = group_cost(rows)
    t0 = 20.0
    best, best_cost = [r[:] for r in rows], cost
    for i in range(iters):
        t = t0 * (1 - i / iters) + 0.01
        d = rng.randrange(DAYS)
        a, b = rng.sample(range(n), 2)
        if rows[a][d] == rows[b][d]:
            continue
        rows[a][d], rows[b][d] = rows[b][d], rows[a][d]
        new = group_cost(rows)
        if new <= cost or rng.random() < math.exp((cost - new) / t):
            cost = new
            if cost < best_cost:
                best, best_cost = [r[:] for r in rows], cost
        else:
            rows[a][d], rows[b][d] = rows[b][d], rows[a][d]
    return best, best_cost


def main():
    rng = random.Random(2026)
    out = {"year": YEAR, "month": MONTH, "hours": HOURS, "demand": DEMAND, "departments": {}}
    for dept, positions in DEMAND.items():
        out["departments"][dept] = {}
        for pos, norms in positions.items():
            n = team_size(norms)
            best_rows, best_cost = None, None
            for attempt in range(3):
                rows = build(norms, n, rng)
                rows, cost = anneal(rows, rng)
                if best_cost is None or cost < best_cost:
                    best_rows, best_cost = rows, cost
            v = sum(violations(r) for r in best_rows)
            hrs = [sum(hours(r[d], d) for d in range(DAYS)) for r in best_rows]
            print(f"{dept:6} {pos:24} n={n:2} cost={best_cost:8.1f} viol={v} "
                  f"hours={min(hrs):.1f}–{max(hrs):.1f}")
            assert v == 0, f"Порушення правил: {pos}"
            out["departments"][dept][pos] = best_rows
    Path(__file__).with_name("schedule.json").write_text(
        json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main()
