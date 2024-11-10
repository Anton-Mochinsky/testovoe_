import pandas as pd
import matplotlib.pyplot as plt

data = pd.read_excel('data.xlsx')

data['receiving_date'] = pd.to_datetime(data['receiving_date'], errors='coerce')

data = data.dropna(subset=['receiving_date'])

def filter_valid_payments(month, year):
    return ((data['receiving_date'].dt.month == month) & (data['receiving_date'].dt.year == year) &
            (data['status'] != 'ПРОСРОЧЕНО'))

if 'sale' in data.columns:
    september_2021_data = data[(data['receiving_date'].dt.month == 9) & (data['receiving_date'].dt.year == 2021)]
    manager_revenue_september = september_2021_data.groupby('sale')['sum'].sum()
    top_manager_september = manager_revenue_september.idxmax()
    top_amount_september = manager_revenue_september.max()
    print(f"Менеджер с наибольшей выручкой в сентябре 2021: {top_manager_september} - {top_amount_september}")
else:
    print("Столбец 'sale' не найден в данных.")

# 1. Общая выручка за июль 2021 по не просроченным платежам
july_2021_data = data[filter_valid_payments(7, 2021)]
total_revenue_july = july_2021_data['sum'].sum()
print(f"Общая выручка за июль 2021: {total_revenue_july}")

# 2. Выручка компании за весь период
data['month_year'] = data['receiving_date'].dt.to_period('M')
revenue_over_time = data.groupby('month_year')['sum'].sum()

# Построение графика
revenue_over_time.plot(kind='line')
plt.title('Выручка компании по месяцам')
plt.xlabel('Месяц')
plt.ylabel('Выручка')
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

# 3. Менеджеры, привлекшие больше всего денежных средств в сентябре 2021
september_2021_data = data[(data['receiving_date'].dt.month == 9) & (data['receiving_date'].dt.year == 2021)]
print("Количество строк в september_2021_data:", september_2021_data.shape[0])

# Проверяем наличие NaN в 'sale'
print("Количество NaN в столбце 'sale':", september_2021_data['sale'].isnull().sum())

# Проверяем, какие уникальные значения есть в столбце 'sale'
print("Уникальные значения в 'sale':", september_2021_data['sale'].unique())

if 'sale' in september_2021_data.columns:
    manager_revenue_september = september_2021_data.groupby('sale')['sum'].sum()
    top_manager_september = manager_revenue_september.idxmax()
    top_amount_september = manager_revenue_september.max()
    print(f"Менеджер с наибольшей выручкой в сентябре 2021: {top_manager_september} - {top_amount_september}")
else:
    print("Столбец 'sale' не найден в сентябре 2021 данных.")

# 4. Преобладающий тип сделок в октябре 2021
october_2021_data = data[(data['receiving_date'].dt.month == 10) & (data['receiving_date'].dt.year == 2021)]
transaction_type_counts = october_2021_data['new/current'].value_counts()
most_common_transaction_type = transaction_type_counts.idxmax()
print(f"Преобладающий тип сделок в октябре 2021: {most_common_transaction_type}")

# 5. Количество оригиналов договора по майским сделкам, полученных в июне 2021
may_2021_data = data[(data['receiving_date'].dt.month == 5) & (data['receiving_date'].dt.year == 2021)]
june_2021_received_orig = may_2021_data[may_2021_data['receiving_date'].dt.month == 6].shape[0]
print(f"Количество оригиналов договора по майским сделкам, полученных в июне 2021: {june_2021_received_orig}")

# Функция для расчета бонусов
def calculate_bonus(row):
    if row['new/current'] == 'новая' and row['status'] == 'ОПЛАЧЕНО' and pd.notna(row['document']):
        return row['sum'] * 0.07
    elif row['new/current'] == 'текущая' and pd.notna(row['document']):
        if row['sum'] > 10000:
            return row['sum'] * 0.05
        else:
            return row['sum'] * 0.03
    return 0

# 6. Остаток бонусов на 01.07.2021
data['bonus'] = data.apply(calculate_bonus, axis=1)
bonuses_july = data[(data['receiving_date'] <= '2021-06-30')].groupby('sale')['bonus'].sum()
print("Остаток бонусов на 01.07.2021:")
print(bonuses_july)
