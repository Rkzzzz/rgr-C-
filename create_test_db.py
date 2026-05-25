#!/usr/bin/env python3
"""
Скрипт для создания тестовой базы данных MS Access с данными о расходах семьи.
"""

import os

try:
    import win32com.client
    USE_COM = True
except ImportError:
    USE_COM = False

def create_access_db_com(db_path):
    """Создание БД через COM (только Windows с установленным MS Access)."""
    access = win32com.client.Dispatch("Access.Application")
    db = access.CurrentDb()
    
    # Создаем новую базу данных
    access.NewCurrentDatabase(db_path)
    
    # Создаем таблицу Expenses
    sql_create_table = """
    CREATE TABLE Expenses (
        ID AUTOINCREMENT PRIMARY KEY,
        ExpenseDate DATE,
        Amount DOUBLE,
        Category TEXT(50),
        FamilyMember TEXT(50)
    )
    """
    db.Execute(sql_create_table)
    
    # Вставляем тестовые данные
    test_data = [
        ("2024-01-15", 5000, "Еда", "Иван"),
        ("2024-01-15", 1500, "Транспорт", "Мария"),
        ("2024-01-16", 3000, "Развлечения", "Иван"),
        ("2024-01-16", 2500, "Еда", "Мария"),
        ("2024-01-17", 800, "Транспорт", "Петр"),
        ("2024-01-17", 4500, "Еда", "Иван"),
        ("2024-01-18", 6000, "Развлечения", "Мария"),
        ("2024-01-18", 1200, "Транспорт", "Петр"),
        ("2024-01-19", 3500, "Еда", "Мария"),
        ("2024-01-19", 2000, "Развлечения", "Петр"),
        ("2024-01-20", 4000, "Еда", "Иван"),
        ("2024-01-20", 1800, "Транспорт", "Мария"),
    ]
    
    for date, amount, category, member in test_data:
        sql_insert = f"""
        INSERT INTO Expenses (ExpenseDate, Amount, Category, FamilyMember) 
        VALUES (#{date}#, {amount}, '{category}', '{member}')
        """
        db.Execute(sql_insert)
    
    db.Close()
    access.Quit()
    print(f"База данных создана: {db_path}")


def create_sqlite_as_alternative(db_path):
    """Создание SQLite БД как альтернативы (кроссплатформенный вариант)."""
    import sqlite3
    
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Создаем таблицу
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Expenses (
            ID INTEGER PRIMARY KEY AUTOINCREMENT,
            ExpenseDate DATE,
            Amount REAL,
            Category TEXT(50),
            FamilyMember TEXT(50)
        )
    ''')
    
    # Вставляем тестовые данные
    test_data = [
        ("2024-01-15", 5000, "Еда", "Иван"),
        ("2024-01-15", 1500, "Транспорт", "Мария"),
        ("2024-01-16", 3000, "Развлечения", "Иван"),
        ("2024-01-16", 2500, "Еда", "Мария"),
        ("2024-01-17", 800, "Транспорт", "Петр"),
        ("2024-01-17", 4500, "Еда", "Иван"),
        ("2024-01-18", 6000, "Развлечения", "Мария"),
        ("2024-01-18", 1200, "Транспорт", "Петр"),
        ("2024-01-19", 3500, "Еда", "Мария"),
        ("2024-01-19", 2000, "Развлечения", "Петр"),
        ("2024-01-20", 4000, "Еда", "Иван"),
        ("2024-01-20", 1800, "Транспорт", "Мария"),
    ]
    
    cursor.executemany(
        'INSERT INTO Expenses (ExpenseDate, Amount, Category, FamilyMember) VALUES (?, ?, ?, ?)',
        test_data
    )
    
    conn.commit()
    conn.close()
    print(f"SQLite база данных создана: {db_path}")


if __name__ == "__main__":
    # Путь к базе данных
    access_db_path = os.path.join(os.path.dirname(__file__), "family_expenses.mdb")
    sqlite_db_path = os.path.join(os.path.dirname(__file__), "family_expenses.sqlite")
    
    if USE_COM:
        try:
            create_access_db_com(access_db_path)
        except Exception as e:
            print(f"Ошибка при создании Access БД: {e}")
            print("Создаю SQLite базу как альтернативу...")
            create_sqlite_as_alternative(sqlite_db_path)
    else:
        print("win32com не доступен (не Windows или нет pywin32)")
        print("Создаю SQLite базу как альтернативу...")
        create_sqlite_as_alternative(sqlite_db_path)
    
    print("\nГотово! База данных с тестовыми данными создана.")
