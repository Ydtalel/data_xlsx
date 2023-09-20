import pandas as pd
import matplotlib.pyplot as plt


# Функция для чтения данных из файла
def read_data(file_path):
    return pd.read_excel(file_path)


# Функция для предварительной обработки данных
def preprocess_data(data):
    data['sum'] = data['sum'].apply(lambda x: str(x).replace(' ', '').replace(',', '.').replace('\xa0', ''))
    data['sum'] = data['sum'].astype(float)
    data['receiving_date'] = pd.to_datetime(data['receiving_date'], format='%d %B %Y', errors='coerce')
    return data


# Функция для задания 1: Вычислить общую выручку за июль 2021
def calculate_total_revenue(data):
    month = data[(data['status'] == "ОПЛАЧЕНО")]
    return round(month['sum'].sum(), 2)


# Функция для задания 2: Построить график изменения выручки
def plot_revenue_time_series(data):
    revenue_time_series = data.groupby('receiving_date')['sum'].sum()
    plt.figure(figsize=(12, 6))
    plt.plot(revenue_time_series.index, revenue_time_series.values, marker='o', linestyle='-')
    plt.title('Изменение выручки компании за период')
    plt.xlabel('Дата')
    plt.ylabel('Выручка')
    plt.grid(True)
    plt.tight_layout()
    plt.show()


# Функция для задания 3: Найти менеджера с наибольшей суммой денежных средств в сентябре 2021
def find_top_manager(data):
    manager_sales = data.groupby('sale')['sum'].sum()
    return manager_sales.idxmax(), manager_sales.max()


# Функция для задания 4: Найти преобладающий тип сделок в октябре 2021
def find_predominant_deal_type(data):
    deal_counts = data['new/current'].value_counts()
    return deal_counts.idxmax()


# Функция для получения фрейма за конкретный месяц
def split_data_frames(data):
    data_frames = []
    start_index = 0

    for idx, row in data.iterrows():
        if pd.isna(row['client_id']):  # Проверяем, является ли client_id пустым значением
            data_frame = data.iloc[start_index:idx]
            data_frames.append(data_frame)
            start_index = idx + 1

    # Добавляем последний фрейм данных (после последнего разделителя)
    if start_index < len(data):
        data_frame = data.iloc[start_index:]
        data_frames.append(data_frame)

    return data_frames


# Функция для задания 5: Сколько оригиналов договора по майским сделкам было получено в июне 2021?
def count_originals_in_month(data, target_month, target_year):
    matching_rows = data[
        (data['receiving_date'].dt.month == target_month) &
        (data['receiving_date'].dt.year == target_year)
        ]
    return matching_rows.shape[0]


# Функция для расчета бонуса за сделки
def calculate_bonus(data, target_month, target_year):
    bonus_remainder = {}
    # Отфильтруем данные, чтобы учесть только сделки, которые были заключены до target_date
    filtered_data = data[
        (data['receiving_date'].dt.month > target_month) &
        (data['receiving_date'].dt.year == target_year)
        ]

    for manager, group in filtered_data.groupby('sale'):
        # Инициализируем бонус текущего менеджера
        bonus = 0.0

        for _, row in group.iterrows():
            status = row['status']
            new_current = row['new/current']
            sales_amount = row['sum']

            # Используем информацию из столбца "document" для определения наличия оригинала
            has_original = "оригинал" in str(row['document']).lower()

            # Проверяем условия задания и рассчитываем бонус
            if new_current == 'новая' and status == 'ОПЛАЧЕНО' and has_original:
                bonus += sales_amount * 0.07  # 7% за новые сделки
            elif new_current == 'текущая' and status != 'ПРОСРОЧЕНО' and has_original:
                if sales_amount > 10000:
                    bonus += sales_amount * 0.05  # 5% за текущие сделки (сумма > 10 тыс.)
                else:
                    bonus += sales_amount * 0.03  # 3% за текущие сделки (сумма <= 10 тыс.)

        # Добавляем бонус текущего менеджера в словарь остатков бонусов
        bonus_remainder[manager] = bonus

    return bonus_remainder


# Основной код
if __name__ == "__main__":
    file_path = "data.xlsx"
    data = read_data(file_path)
    data = preprocess_data(data)
    frame_list = split_data_frames(data)

    june = frame_list[2]
    july = frame_list[3]
    september = frame_list[5]
    october = frame_list[6]
    print(f'Общая выручка за июль: {calculate_total_revenue(july)}')
    manager, sales_sum = find_top_manager(september)
    print(f'Лучший менеджер - {manager}, с объемом продаж {sales_sum}')
    print(f'Тип сделок "{find_predominant_deal_type(october)}" был преобладающим в октябре 2021')
    print(f'Количество договоров по майским сделкам: {count_originals_in_month(june, 5, 2021)} ')

    # Рассчитываем остаток бонусов на 01.07.2021
    bonus_remainder_on_july_1 = calculate_bonus(june, 6, 2021)

    print("Остаток бонусов на 01.07.2021:")
    for manager, bonus_remainder in bonus_remainder_on_july_1.items():
        print(f"{manager}: {bonus_remainder:.2f}")
