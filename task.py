import pandas as pd

""""1. Загрузка файлов"""

garages_df = pd.read_excel('arenda.xlsx')
statement_df = pd.read_excel('print 2.xlsx')

"""2. Подготовка данных выписки"""

# Считываем выписку, пропуская первые 18 строк до данных
statement_data = pd.read_excel(
    'print 2.xlsx',
    skiprows=18,
    usecols="A:E",
    names=["Дата операции", "Время", "Empty", "Описание", "Сумма"]
)

# Удаляем строки без суммы или с заголовками
statement_data = statement_data.dropna(subset=["Дата операции", "Сумма"])
statement_data = statement_data[statement_data["Описание"].notna()]
statement_data = statement_data[~statement_data["Сумма"].astype(str).str.contains("СУММА", na=False)]

# Преобразуем сумму к float
statement_data["Сумма"] = (
    statement_data["Сумма"]
    .astype(str)
    .str.replace(" ", "")
    .str.replace("+", "")
    .str.replace(",", ".")
    .astype(float)
)

"""3. Подготовка таблицы гаражей"""

garages_df["Сумма"] = garages_df["Сумма"].astype(float)

"""4. Объединение таблиц по сумме"""

result = pd.merge(
    garages_df,
    statement_data[["Сумма", "Дата операции"]],
    on="Сумма",
    how="left"
)

# Добавляем колонку 'Оплачено' со значениями 'ДА' или 'НЕТ'
result["Оплачено"] = result["Дата операции"].notna().apply(lambda x: "ДА" if x else "НЕТ")

"""5. Сохраняем результат"""

result.to_excel("garages_checked.xlsx", index=False)

print("Сверка завершена. Результат сохранён в 'garages_checked.xlsx'.")
