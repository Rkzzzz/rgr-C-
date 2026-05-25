"""
Скрипт для создания тестовой базы данных MS Access с данными о расходах семьи.
Требуется установленный Python и пакет pyodbc + access (или используйте SQLite версию).
"""

import sqlite3
from datetime import datetime, timedelta
import random

def create_test_database():
    """Создает тестовую SQLite базу данных (альтернатива MS Access для тестирования)"""
    
    conn = sqlite3.connect('family_expenses.sqlite')
    cursor = conn.cursor()
    
    # Создание таблицы Expenses
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Expenses (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            ExpenseDate DATE NOT NULL,
            Amount DECIMAL(10,2) NOT NULL,
            Category TEXT NOT NULL,
            MemberName TEXT NOT NULL
        )
    ''')
    
    # Очистка существующих данных
    cursor.execute('DELETE FROM Expenses')
    
    # Данные для генерации
    categories = ['Еда', 'Транспорт', 'Развлечения', 'Коммунальные услуги', 
                  'Одежда', 'Здоровье', 'Образование', 'Прочее']
    members = ['Александр', 'Мария', 'Дмитрий', 'Елена', 'Анна']
    
    # Генерация тестовых данных за последние 30 дней
    base_date = datetime.now()
    expenses_data = []
    
    for i in range(50):  # 50 записей
        date = base_date - timedelta(days=random.randint(0, 30))
        amount = round(random.uniform(100, 15000), 2)
        category = random.choice(categories)
        member = random.choice(members)
        
        expenses_data.append((date.strftime('%Y-%m-%d'), amount, category, member))
    
    # Вставка данных
    cursor.executemany(
        'INSERT INTO Expenses (ExpenseDate, Amount, Category, MemberName) VALUES (?, ?, ?, ?)',
        expenses_data
    )
    
    conn.commit()
    conn.close()
    
    print("База данных family_expenses.sqlite успешно создана!")
    print(f"Добавлено {len(expenses_data)} записей о расходах.")
    print("\nСтруктура таблицы:")
    print("- Id: INTEGER (первичный ключ)")
    print("- ExpenseDate: DATE (дата расхода)")
    print("- Amount: DECIMAL (сумма)")
    print("- Category: TEXT (категория: Еда, Транспорт, Развлечения и т.д.)")
    print("- MemberName: TEXT (имя члена семьи)")

if __name__ == '__main__':
    create_test_database()
