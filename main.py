import csv
from collections import defaultdict
import openpyxl
from openpyxl.utils import get_column_letter

def read_data(file_name):
    data = []
    with open(file_name, newline='') as csvfile:
        reader = csv.reader(csvfile, delimiter='|')
        next(reader)  # Скип заголовок
        for row in reader:
            order_number, order_date, product_name, category, quantity, price_per_unit, total_cost = [item.strip() for item in row]
            data.append({
                'order_number': order_number,
                'order_date': order_date,
                'product_name': product_name,
                'category': category,
                'quantity': int(quantity),
                'price_per_unit': float(price_per_unit),
                'total_cost': float(total_cost)
            })
    return data

def calculate_total_revenue(data):
    total_revenue = sum(item['total_cost'] for item in data)
    return total_revenue

def find_popular_product(data):
    products_sold = defaultdict(int)
    for item in data:
        products_sold[item['product_name']] += item['quantity']
    most_sold_product = max(products_sold, key=products_sold.get)
    return most_sold_product

def find_most_revenue_product(data):
    revenue_per_product = defaultdict(float)
    for item in data:
        revenue_per_product[item['product_name']] += item['total_cost']
    most_revenue_product = max(revenue_per_product, key=revenue_per_product.get)
    return most_revenue_product

def print_sales_info(data):
    total_revenue = calculate_total_revenue(data)
    print("Общая выручка магазина:", total_revenue)
    print('-------------------------------------------------------------')

    most_sold_product = find_popular_product(data)
    print("Самая продаваемая животинка:", most_sold_product)
    print('-------------------------------------------------------------')

    most_revenue_product = find_most_revenue_product(data)
    print("Животинка, принесшая наибольшую выручку:", most_revenue_product)
    print('-------------------------------------------------------------')

    print("Информация о продажах каждой животинки:")
    products_sold = defaultdict(int)
    for item in data:
        products_sold[item['product_name']] += item['quantity']
    for product, quantity in products_sold.items():
        revenue = sum(item['total_cost'] for item in data if item['product_name'] == product)
        revenue_share = (revenue / total_revenue) * 100
        print(f"{product}: Продано {quantity} единиц, Выручка: {revenue}, Доля в общей выручке: {revenue_share:.2f}%")

def save_to_excel(data, filename='output.xlsx'):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Our_best_of_the_best_profit"

    headers = ["Тески животных", "Сколько продали", "Выручка", "Доля в общей выручке в %"]
    ws.append(headers)

    total_revenue = calculate_total_revenue(data)
    products_sold = defaultdict(int)
    for item in data:
        products_sold[item['product_name']] += item['quantity']

    for product, quantity in products_sold.items():
        revenue = sum(item['total_cost'] for item in data if item['product_name'] == product)
        revenue_share = (revenue / total_revenue) * 100
        ws.append([product, quantity, revenue, f"{revenue_share:.2f}"])

    for col in ws.columns:
        max_length = 0
        col_name = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[col_name].width = adjusted_width

    wb.save(filename)

file_name = 'sales_data.csv'
try:
    data = read_data(file_name)
    print_sales_info(data)
    save_to_excel(data)
    print(f"Данные сохранены в файл {file_name}")
except FileNotFoundError:
    print("Файл не найден")
except Exception as e:
    print("Произошла ошибка:", e)
