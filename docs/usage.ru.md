# FlowXL

**Язык:** [English](usage.md) · Русский

Предсказуемая запись Excel поверх openpyxl.

Громкая ошибка вместо тихой порчи данных. Связный API для создания и правки `.xlsx`.

## Установка

```bash
pip install -e .
# Необязательные extras.
pip install -e '.[polars]'
pip install -e '.[pandas]'
```

Python 3.12+. Единственная обязательная зависимость — openpyxl.

## Что импортировать

Всё публичное живёт в `flowxl`. Нет в этом списке — в патче могут переименовать.

```python
from flowxl import (
    Workbook,  # Создать, открыть или потоком записать книгу.
    Sheet,  # Лист со случайным доступом из Workbook.sheet.
    StreamingWorkbook,  # Книга только для записи из Workbook.stream.
    StreamingSheet,
    WorkbookSettings,  # Тема, политика имён листов, потолок autofit.
    CellRef,  # Одна ячейка, с единицы. Разбор через parse('B12').
    CellRange,  # Прямоугольник. Разбор через parse('A1:C10').
    ORIGIN,  # Ячейка A1.
    KEEP,  # Значение не трогать.
    StyleSpec,  # Чертёж именованного стиля.
    FontSpec,
    FillSpec,
    Color,
    DefaultTheme,
    Theme,
    to_excel_formula,  # Русская формула в en-US текст в xlsx.
    FlowxlError,  # Поймать всё, что бросила библиотека.
    UsageError,  # Вызов неправильный.
)
```

## Создать, открыть, сохранить

```python
from flowxl import Workbook

with Workbook.new() as wb:  # Пустой документ, листов ещё нет.
    wb.sheet('Sales')
    wb.save('sales.xlsx')  # Save вызываете сами. Блок with только закрывает.

with Workbook.open('sales.xlsx') as wb:
    print(wb.sheet_names)
    wb.save('sales.xlsx')  # Save можно звать несколько раз.

raw = Workbook.new().to_bytes()  # Тот же документ в памяти.
wb = Workbook.from_bytes(raw)
```

`with` книгу **закрывает**. Он **не** сохраняет.

| Вызов | Что делает |
|---|---|
| `Workbook.new()` | Пустая книга |
| `Workbook.open(path, formulas=True)` | Открыть файл. `formulas=False` читает кэш значений |
| `Workbook.from_bytes(data)` | Открыть из памяти |
| `Workbook.stream()` | Книга только для записи, для больших таблиц |
| `wb.sheet(name)` | Взять или создать лист |
| `wb.has_sheet(name)` | Есть ли лист |
| `wb.remove(name)` | Удалить лист |
| `wb.sheet_names` | Имена в порядке документа |
| `wb.save(path)` | Записать на диск; книга остаётся рабочей |
| `wb.to_bytes()` | Записать в память |
| `wb.close()` | Отпустить ресурсы |

## Ячейки

Координаты с единицы. `CellRef.parse('B12')` — строка 12, колонка 2.

```python
from datetime import date
from flowxl import CellRef, FontSpec, StyleSpec, Workbook

bold = StyleSpec(name='bold', font=FontSpec(bold=True))

with Workbook.new() as wb:
    sheet = wb.sheet('Report')
    sheet.cell(CellRef.parse('A1'), 'Hello')
    sheet.cell(CellRef.parse('B1'), 42)
    sheet.cell(CellRef.parse('C1'), date(2024, 1, 2))
    sheet.cell(CellRef.parse('A1'), style=bold)  # Без значения: только стиль.
    sheet.cell(CellRef.parse('B1'), None)  # None очищает ячейку.
    print(sheet.read(CellRef.parse('A1')))
    wb.save('report.xlsx')
```

| Вызов | Что делает |
|---|---|
| `sheet.cell(ref, value=KEEP, style=None)` | Записать значение и/или стиль |
| `sheet.row(at, values, style=None)` | Слева направо |
| `sheet.column(at, values, style=None)` | Сверху вниз |
| `sheet.read(ref)` | Текущее хранимое значение |
| `sheet.merge(area)` | Объединить прямоугольник |

`KEEP` значит «значение не трогать». `None` значит «очистить ячейку». Это разные вещи.

## Формулы

Библиотека кладёт формулу как текст. Excel считает её, когда файл открывают.

xlsx всегда хранит **английские** имена и разделители **en-US**. Русский Excel только *показывает* `СУММ` и `;`. FlowXL переводит при записи:

```python
sheet.formula(CellRef.parse('D1'), '=СУММ(B1;C1)')  # В файле =SUM(B1,C1).
sheet.formula(CellRef.parse('E1'), '=A1+1,5')  # В файле =A1+1.5.
sheet.formula(CellRef.parse('F1'), '=SUM(B1,C1)')  # Уже английский: как есть.
```

`cell(ref, '=СРЗНАЧ(A1:A10)')` переводится так же: любая строка с `=` идёт через `to_excel_formula`.

| Вызов | Что делает |
|---|---|
| `sheet.formula(ref, expression, style=None)` | Записать формулу. Должна начинаться с `=` |
| `to_excel_formula(expression)` | Перевести без книги |

Неизвестная кириллическая функция — `UsageError`. Текст в кавычках и имена листов не трогаем.

## Строки, колонки, диапазоны

`CellRange` — прямоугольник для merge, autofilter и выпадающего списка.

```python
from flowxl import CellRange, CellRef, Workbook

with Workbook.new() as wb:
    sheet = wb.sheet('Report')
    sheet.cell(CellRef.parse('A1'), 'Q1')
    sheet.merge(CellRange.parse('A1:C1'))
    sheet.row(CellRef.parse('A3'), ['Name', 'Qty', 'Total'])
    sheet.column(CellRef.parse('A4'), ['Ann', 'Bob'])
    sheet.widths({'A': 16, 'B': 8, 'C': 10})
    sheet.heights({1: 24})
    sheet.freeze(CellRef.parse('A4'))  # Первая незамороженная ячейка.
    sheet.pin_header()  # Закрепить первую строку.
    sheet.autofilter()  # Накрывает всё, что записали через FlowXL.
    sheet.autofit()
    wb.save('report.xlsx')
```

| Вызов | Что делает |
|---|---|
| `sheet.freeze(ref)` | Закрепить всё выше и левее `ref` |
| `sheet.pin_header(rows=1)` | Закрепить верхние строки |
| `sheet.autofilter(area=None)` | Фильтр. Без `area` — по учтённым границам |
| `sheet.widths({1: 12})` / `sheet.widths({'A': 12})` | Ширины колонок |
| `sheet.heights({1: 18})` | Высоты строк |
| `sheet.autofit(max_width=None, only=None)` | Подогнать колонки под записанное |

После `sheet.raw_openpyxl()` границам больше нельзя верить. Передайте `area=` в `autofilter`.

## Таблицы

`frame()` пишет таблицу и оформляет колонки по логическому типу. Список словарей или DataFrame polars / pandas.

```python
from datetime import date
from flowxl import CellRef, Workbook

records = [
    {'Name': 'Ann', 'Amount': 10, 'Day': date(2024, 1, 2)},
    {'Name': 'Bob', 'Amount': 20, 'Day': date(2024, 1, 3)},
]

with Workbook.new() as wb:
    wb.sheet('Sales').frame(records).pin_header().autofilter().autofit()
    wb.sheet('Offset').frame(records, at=CellRef.parse('B2'))
    wb.save('sales.xlsx')
```

| Вызов | Что делает |
|---|---|
| `sheet.frame(data, at=ORIGIN, header=True, styles=None)` | Записать таблицу. `styles` это `{имя_колонки: StyleSpec}` |

## Стили

Даты, числа, проценты и валюта берут встроенные стили Excel (`Normal`, `Comma [0]`, `Comma`, `Percent`, `Currency`). Свой именованный стиль по умолчанию только у шапки: `sw_header`.

У своего `StyleSpec` обязательно **имя**.

```python
from flowxl import (
    CellRef,
    Color,
    DefaultTheme,
    FillSpec,
    FontSpec,
    StyleSpec,
    Workbook,
    WorkbookSettings,
)

title = StyleSpec(
    name='title',
    font=FontSpec(size=16, bold=True, color=Color.parse('#FFFFFF')),
    fill=FillSpec.solid(Color.parse('#007CD6')),
)
percent = StyleSpec(name='pct', number_format='0.0%')
records = [{'Name': 'Ann', 'Amount': 0.12}, {'Name': 'Bob', 'Amount': 0.88}]

with Workbook.new() as wb:
    sheet = wb.sheet('Sales')
    sheet.cell(CellRef.parse('A1'), 'Q1', style=title)
    sheet.frame(records, at=CellRef.parse('A3'), styles={'Amount': percent})
    wb.save('sales.xlsx')

money = StyleSpec(name='usd', number_format='$#,##0.00')
settings = WorkbookSettings(theme=DefaultTheme().with_overrides(float=money))
with Workbook.new(settings=settings) as wb:
    wb.sheet('Sales').frame(records)
    wb.save('money.xlsx')
```

Чтобы взять встроенный стиль Excel и не создавать новый named style:

```python
StyleSpec(name='my_percent', builtin='Percent')
StyleSpec(name='my_date', builtin='Normal', number_format='mm-dd-yy')
```

## Выпадающий список

```python
sheet.dropdown(CellRef.parse('A4'), ['Ann', 'Bob'])
sheet.dropdown(CellRange.parse('B1:B10'), CellRange.parse('A1:A3'))  # Варианты уже на листе.
```

Запятая или кавычка внутри варианта разорвёт список Excel. Положите такие варианты на лист и передайте `CellRange`. Одну строку `'Yes'` не принимаем: Python прошёл бы по буквам.

| Вызов | Что делает |
|---|---|
| `sheet.dropdown(area, options, allow_blank=True)` | Список выбора на ячейке или диапазоне |

## Защита листа

```python
sheet.protect()  # Закрыть лист.
sheet.protect('secret')  # Закрыть с паролем.
sheet.unprotect()
```

По умолчанию каждая ячейка locked. После `protect()` править нельзя, пока стиль не скажет `ProtectionSpec(locked=False)`.

| Вызов | Что делает |
|---|---|
| `sheet.protect(password=None)` | Включить защиту листа |
| `sheet.unprotect()` | Снять защиту |

## Поток

Только запись, для больших таблиц. Нет `cell()`, `read()`, `merge()`. Ширины колонок и закреплённую шапку задают при создании листа. `protect` и `dropdown` работают: они украшают лист, не пишут ячейку.

```python
from flowxl import Workbook

rows = [{'Name': 'Ann', 'Amount': 10}, {'Name': 'Bob', 'Amount': 20}]

with Workbook.stream() as wb:
    sheet = wb.sheet('Events', pin_header=True, widths={'A': 24, 'B': 12})
    sheet.frame(rows)
    wb.save('events.xlsx')
```

Потоковые `save` / `to_bytes` книгу **закрывают**. Чтобы писать снова, создайте новую.

## Ошибки

Ловите `FlowxlError`, если нужно любое падение FlowXL.

- `UsageError` — виноват вызов (плохой адрес, формула без `=`, неизвестный тип). Чините программу.
- `BackendError` — файл, диск или openpyxl. Попробуйте другой путь.

Сообщения английские. В тексте всегда есть что случилось и что попробовать.

## Чего библиотека не делает

См. [scope.md](scope.md). Не считает формулы, не читает `.xls` / `.csv`, не заменяет openpyxl.
