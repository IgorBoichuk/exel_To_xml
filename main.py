# import pandas as pd
# from datetime import datetime
# import xml.etree.ElementTree as ET
# from ftplib import FTP
# import os

# # Завантажуємо дані з Excel
# file_path = 'input_price.xlsx'
# df = pd.read_excel(file_path)

# # Створення основи для XML
# current_datetime = datetime.now().strftime('%Y-%m-%d %H:%M')
# yml_catalog = ET.Element('yml_catalog', date=current_datetime)
# shop = ET.SubElement(yml_catalog, 'shop')

# # Додаємо основну інформацію про магазин
# ET.SubElement(shop, 'name').text = 'Klipster'
# ET.SubElement(shop, 'company').text = 'Klipster'

# # Додаємо категорії (перебираємо унікальні категорії з Excel)
# categories = ET.SubElement(shop, 'categories')
# category_ids = {}  # Словник для збереження категорій з унікальними ID
# category_id_counter = 1  # Лічильник для унікальних ID категорій

# for category in df['category'].unique():
#     # Додаємо категорію тільки, якщо її ще немає в словнику
#     if category not in category_ids:
#         category_ids[category] = category_id_counter
#         ET.SubElement(categories, 'category', id=str(category_id_counter)).text = category
#         category_id_counter += 1

# # Додаємо пропозиції
# offers = ET.SubElement(shop, 'offers')

# # Для кожного товару в Excel
# for index, row in df.iterrows():
#     offer = ET.SubElement(offers, 'offer', id=str(row['article']), available='true')
    
#     # Ціна
#     ET.SubElement(offer, 'price').text = str(row['price'])
    
#     # Категорія (відповідно до ID)
#     category_id = category_ids.get(row['category'], None)
#     if category_id:
#         ET.SubElement(offer, 'categoryId').text = str(category_id)
    
#     # Фото
#     ET.SubElement(offer, 'picture').text = str(row['picture1'])
    
#     # Вендор
#     ET.SubElement(offer, 'vendor').text = str(row['vendor'])
    
#     # Артикул
#     ET.SubElement(offer, 'article').text = str(row['article'])
    
#     # Кількість на складі - наявність
#     stock_quantity = row['available']
#     if stock_quantity > 10:
#         ET.SubElement(offer, 'stock_quantity').text = 'true'
#     else:
#         ET.SubElement(offer, 'stock_quantity').text = 'false'
    
#     # Назва російською
#     ET.SubElement(offer, 'name').text = str(row['name'])
    
#     # Назва українською
#     ET.SubElement(offer, 'name_ua').text = str(row['name_ua'])
    
#     # Опис російською
#     description_text = row['description']
#     description = ET.SubElement(offer, 'description')
#     if isinstance(description_text, str) and not pd.isna(description_text):
#         description.text = f"<![CDATA[{description_text[:500]}]]>"  # Лімітуємо опис до 500 символів
#     else:
#         description.text = ""
    
#     # Опис українською
#     description_ua_text = row['description_ua']
#     description_ua = ET.SubElement(offer, 'description_ua')
#     if isinstance(description_ua_text, str) and not pd.isna(description_ua_text):
#         description_ua.text = f"<![CDATA[{description_ua_text[:500]}]]>"  # Лімітуємо опис до 500 символів
#     else:
#         description_ua.text = ""

#     # Додаткові параметри
#     for col in ['type', 'color', 'instalation_place','auto_brand','min_cart_quantity']:
#         param = ET.SubElement(offer, 'param', name=col)
#         param.text = str(row[col])

# # Створення XML дерева
# tree = ET.ElementTree(yml_catalog)
# output_file = "price.xml"

# # Записуємо XML у файл
# with open(output_file, "wb") as f:
#     tree.write(f, encoding="UTF-8", xml_declaration=True)

# print("XML файл створено!")

# # Завантаження на FTP
# ftp_host = "109.94.209.117"
# ftp_port = 21
# ftp_user = "ihorbo01"
# ftp_pass = "Gross37038263"
# ftp_target_dir = "/"  # Замініть на потрібну папку на сервері

# try:
#     with FTP() as ftp:
#         # Підключення до FTP
#         ftp.connect(ftp_host, ftp_port)
#         ftp.login(ftp_user, ftp_pass)
#         print("Підключено до FTP")

#         # Перехід до цільової папки
#         if ftp_target_dir:
#             ftp.cwd(ftp_target_dir)
#             print(f"Папка змінена на: {ftp_target_dir}")

#         # Відкриваємо файл і завантажуємо його
#         with open(output_file, "rb") as file:
#             ftp.storbinary(f"STOR {output_file}", file)
#             print(f"Файл '{output_file}' завантажено на FTP сервер")
# except Exception as e:
#     print(f"Помилка FTP: {e}")
import pandas as pd
from datetime import datetime
import xml.etree.ElementTree as ET

# Завантаження даних з Excel
file_path = 'input_price.xlsx'
df = pd.read_excel(file_path)

# Створення основи для XML
current_datetime = datetime.now().strftime('%Y-%m-%d %H:%M')
yml_catalog = ET.Element('yml_catalog', date=current_datetime)
shop = ET.SubElement(yml_catalog, 'shop')

# Додавання основної інформації про магазин
ET.SubElement(shop, 'name').text = 'Klipster'
ET.SubElement(shop, 'company').text = 'Klipster'

# Додавання категорій
categories = ET.SubElement(shop, 'categories')
category_ids = {}
category_id_counter = 1

for category in df['Категорія'].unique():
    if category not in category_ids:
        category_ids[category] = category_id_counter
        ET.SubElement(categories, 'category', id=str(category_id_counter)).text = category
        category_id_counter += 1

# Додавання пропозицій
offers = ET.SubElement(shop, 'offers')

for _, row in df.iterrows():
    offer = ET.SubElement(offers, 'offer', id=str(row['Код_товару']), available='true')

    # Основні атрибути
    ET.SubElement(offer, 'price').text = str(row['Ціна'])
    category_id = category_ids.get(row['Категорія'])
    if category_id:
        ET.SubElement(offer, 'categoryId').text = str(category_id)
    ET.SubElement(offer, 'picture').text = str(row['Посилання_зображення'])
    ET.SubElement(offer, 'vendor').text = str(row['Виробник'])
    ET.SubElement(offer, 'name').text = str(row['Назва_позиції'])
    ET.SubElement(offer, 'name_ua').text = str(row['Назва_позиції_укр'])

    # Опис
    description_text = row['Опис']
    description = ET.SubElement(offer, 'description')
    description.text = f"<![CDATA[{str(description_text)[:500]}]]>" if isinstance(description_text, str) else ""

    description_ua_text = row['Опис_укр']
    description_ua = ET.SubElement(offer, 'description_ua')
    description_ua.text = f"<![CDATA[{str(description_ua_text)[:500]}]]>" if isinstance(description_ua_text, str) else ""

    # Розміри та вага
    weight = row['Вага,кг']
    ET.SubElement(offer, 'weight').text = "" if pd.isna(weight) else str(weight)

    width = row['Ширина,см']
    ET.SubElement(offer, 'width').text = "" if pd.isna(width) else str(width)

    height = row['Висота,см']
    ET.SubElement(offer, 'height').text = "" if pd.isna(height) else str(height)

    length = row['Довжина,см']
    ET.SubElement(offer, 'length').text = "" if pd.isna(length) else str(length)

    # Характеристики починаються з 21 стовпця
    for i in range(20, len(df.columns), 3):  # Починаємо з 21-го стовпця (0-based index)
        # Перевірка на наявність достатньої кількості стовпців
        if i + 2 < len(df.columns):
            name_col = df.columns[i]  # Назва характеристики
            unit_col = df.columns[i + 1]  # Одиниця виміру
            value_col = df.columns[i + 2]  # Значення характеристики

            # Перевіряємо, чи є значення в "Назва_Характеристики"
            if pd.notna(row[name_col]) and row[name_col] != "":  
                param = ET.SubElement(offer, 'param', name=str(row[name_col]))

                # Якщо є "Одиниця_виміру_Характеристики", додаємо її як атрибут
                if pd.notna(row[unit_col]) and row[unit_col] != "":
                    param.set('unit', str(row[unit_col]))  # Перетворення в рядок

                # Якщо є "Значення_Характеристики", додаємо його як текст
                if pd.notna(row[value_col]) and row[value_col] != "":
                    param.text = str(row[value_col])  # Перетворення всіх значень на рядок
                        
# Створення XML дерева
output_file = "price.xml"
tree = ET.ElementTree(yml_catalog)

# Запис у файл
with open(output_file, "wb") as f:
    tree.write(f, encoding="UTF-8", xml_declaration=True)



print("XML файл створено!")
import os
print("Поточна робоча директорія:", os.getcwd())


# Завантаження на FTP
# ftp_host = "109.94.209.117"
# ftp_port = 21
# ftp_user = "ihorbo01"
# ftp_pass = "Gross37038263"
# ftp_target_dir = "/"

# try:
#     with FTP() as ftp:
#         ftp.connect(ftp_host, ftp_port)
#         ftp.login(ftp_user, ftp_pass)
#         print("Підключено до FTP")

#         if ftp_target_dir:
#             ftp.cwd(ftp_target_dir)
#             print(f"Папка змінена на: {ftp_target_dir}")

#         with open(output_file, "rb") as file:
#             ftp.storbinary(f"STOR {output_file}", file)
#             print(f"Файл '{output_file}' завантажено на FTP сервер")
# except Exception as e:
#     print(f"Помилка FTP: {e}")
