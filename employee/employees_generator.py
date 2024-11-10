from faker import Faker
import csv
import random
from datetime import datetime

# Створюємо екземпляр Faker з українською локалізацією
fake = Faker('uk_UA')

# Визначаємо можливі по батькові для чоловіків і жінок
male_middle_names = [
    "Олександрович", "Миколайович", "Іванович", "Сергійович", "Петрович",
    "Володимирович", "Васильович", "Андрійович", "Юрійович", "Григорович",
    "Романович", "Вікторович", "Богданович", "Олегович", "Михайлович"
]

female_middle_names = [
    "Олександрівна", "Миколаївна", "Іванівна", "Сергіївна", "Петрівна",
    "Володимирівна", "Василівна", "Андріївна", "Юріївна", "Григорівна",
    "Романівна", "Вікторівна", "Богданівна", "Олегівна", "Михайлівна"
]

# Встановлюємо параметри для генерації
num_records = 2000
male_percentage = 0.6
female_percentage = 0.4

# Розраховуємо кількість чоловіків і жінок
num_males = int(num_records * male_percentage)
num_females = num_records - num_males

# Генеруємо дані та записуємо в CSV з кодуванням UTF-8 з BOM
with open('employees.csv', 'w', newline='', encoding='utf-8-sig') as csvfile:
    fieldnames = ["Прізвище", "Ім'я", "По батькові", "Стать", "Дата народження", 
                  "Посада", "Місто проживання", "Адреса проживання", "Телефон", "Email"]
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()
    
    # Функція для генерації дати народження
    def generate_birthdate():
        start_date = datetime(1938, 1, 1)
        end_date = datetime(2008, 12, 31)
        return fake.date_between_dates(date_start=start_date, date_end=end_date)
    
    # Генерація чоловічих записів
    for _ in range(num_males):
        first_name = fake.first_name_male()
        last_name = fake.last_name_male()
        middle_name = random.choice(male_middle_names)
        writer.writerow({
            "Прізвище": last_name,
            "Ім'я": first_name,
            "По батькові": middle_name,
            "Стать": "Чоловік",
            "Дата народження": generate_birthdate().strftime('%d.%m.%Y'),
            "Посада": fake.job(),
            "Місто проживання": fake.city(),
            "Адреса проживання": fake.address(),
            "Телефон": fake.phone_number(),
            "Email": fake.email()
        })
    
    # Генерація жіночих записів
    for _ in range(num_females):
        first_name = fake.first_name_female()
        last_name = fake.last_name_female()
        middle_name = random.choice(female_middle_names)
        writer.writerow({
            "Прізвище": last_name,
            "Ім'я": first_name,
            "По батькові": middle_name,
            "Стать": "Жінка",
            "Дата народження": generate_birthdate().strftime('%d.%m.%Y'),
            "Посада": fake.job(),
            "Місто проживання": fake.city(),
            "Адреса проживання": fake.address(),
            "Телефон": fake.phone_number(),
            "Email": fake.email()
        })

print("Дані успішно згенеровані та збережені у файл employees.csv.")
