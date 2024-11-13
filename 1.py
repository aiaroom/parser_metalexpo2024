import requests
from bs4 import BeautifulSoup
import pytesseract
from PIL import Image
import pandas as pd
import io
import urllib.parse
import re

pytesseract.pytesseract.tesseract_cmd = r"C:\Users\curi\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"

# Базовый URL для участников
base_url = "https://www.metal-expo.ru/ru/exhibition/142/"

# Пустой список для хранения информации о компаниях
companies = []

# Диапазон ID участников
for participant_id in range(70120, 70821):  # Настройте диапазон ID по необходимости
    url = f"{base_url}{participant_id}"
    response = requests.get(url)

    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Извлечение названия компании
        name_tag = soup.find('h2')  # Замените на фактический селектор
        company_name = name_tag.text.strip() if name_tag else "Не указано"
        
        # Извлечение страны компании
        country_tag_dl = soup.find('dl', class_='dl-horizontal pull-left')  # Замените на фактический селектор страны
        adr_dd = country_tag_dl.find('dd',itemprop = 'address')
        country = adr_dd.text.strip() if adr_dd else "Не указано"
        
        # Поиск секции <dl> и элементов <dd> с адресом электронной почты
        dl_tag = soup.find('dl', class_='dl-horizontal pull-left')
        email = "Не указано"
        # print(dl_tag)

        if dl_tag:
            image_tag = dl_tag.find('img')

            if image_tag:
                image_link = 'https://www.metal-expo.ru'+image_tag['src']             
                img_response = requests.get(image_link)
                if img_response.status_code == 200:
                    img = Image.open(io.BytesIO(img_response.content))
                    email = pytesseract.image_to_string(img, lang="eng").strip()
                else:
                    print(f"Не удалось загрузить изображение для {company_name}")
        
        # Добавление данных в список
        companies.append({
            "Название компании": company_name,
            "Почта": email,
            "Страна": country
        })
    else:
        print(f"Не удалось получить данные для участника ID {participant_id}")

# Создание Excel файла
df = pd.DataFrame(companies)
df.to_excel("companies_data.xlsx", index=False)
print("Данные сохранены в файл companies_data.xlsx")