#!/usr/bin/env python3
"""
Приложение для работы с базой данных расходов семьи.
Поддерживает MS Access (.mdb, .accdb) и SQLite базы данных.
Отображает данные в виде круговой диаграммы с группировкой по:
- дате
- сумме
- категории траты
- имени члена семьи
"""

import os
import sys
from datetime import datetime


def connect_to_database(db_path):
    """
    Подключение к базе данных.
    Поддерживает MS Access (через ODBC) и SQLite.
    """
    ext = os.path.splitext(db_path)[1].lower()
    
    if ext in ['.mdb', '.accdb']:
        # Подключение к MS Access через ODBC
        try:
            import pyodbc
            # Строка подключения зависит от драйвера и платформы
            # Для Windows:
            if sys.platform == 'win32':
                conn_str = (
                    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                    f'DBQ={os.path.abspath(db_path)};'
                )
            else:
                # Для Linux/Mac требуется mdbtools или другой драйвер
                conn_str = (
                    r'DRIVER=MDBTools;'
                    f'DATABASE={os.path.abspath(db_path)};'
                )
            conn = pyodbc.connect(conn_str)
            return conn, 'access'
        except Exception as e:
            print(f"Ошибка подключения к Access: {e}")
            print("Попробуйте использовать SQLite базу данных.")
            return None, None
    
    elif ext in ['.sqlite', '.db', '.sqlite3']:
        # Подключение к SQLite
        import sqlite3
        conn = sqlite3.connect(db_path)
        return conn, 'sqlite'
    
    else:
        print(f"Неподдерживаемый формат базы данных: {ext}")
        return None, None


def fetch_expenses(conn, db_type):
    """
    Получение данных о расходах из базы данных.
    Возвращает список кортежей: (дата, сумма, категория, член_семьи)
    """
    cursor = conn.cursor()
    
    if db_type == 'access':
        query = """
        SELECT ExpenseDate, Amount, Category, FamilyMember 
        FROM Expenses 
        ORDER BY ExpenseDate
        """
    else:  # sqlite
        query = """
        SELECT ExpenseDate, Amount, Category, FamilyMember 
        FROM Expenses 
        ORDER BY ExpenseDate
        """
    
    cursor.execute(query)
    rows = cursor.fetchall()
    return rows


def group_by_category(data):
    """Группировка расходов по категории."""
    grouped = {}
    for date, amount, category, member in data:
        if category not in grouped:
            grouped[category] = 0
        grouped[category] += amount
    return grouped


def group_by_member(data):
    """Группировка расходов по члену семьи."""
    grouped = {}
    for date, amount, category, member in data:
        if member not in grouped:
            grouped[member] = 0
        grouped[member] += amount
    return grouped


def group_by_date(data):
    """Группировка расходов по дате."""
    grouped = {}
    for date, amount, category, member in data:
        date_str = str(date) if date else "Не указано"
        if date_str not in grouped:
            grouped[date_str] = 0
        grouped[date_str] += amount
    return grouped


def create_pie_chart(grouped_data, title, filename=None):
    """
    Создание круговой диаграммы.
    grouped_data: словарь {метка: значение}
    title: заголовок диаграммы
    filename: имя файла для сохранения (если None - показать на экране)
    """
    import matplotlib.pyplot as plt
    import matplotlib
    
    # Настройка шрифтов для поддержки кириллицы
    matplotlib.rcParams['axes.unicode_minus'] = False
    
    # Попытка использовать доступные шрифты с поддержкой кириллицы
    available_fonts = ['DejaVu Sans', 'Liberation Sans', 'FreeSans', 'sans-serif']
    matplotlib.rcParams['font.family'] = available_fonts
    
    if not grouped_data:
        print("Нет данных для отображения!")
        return
    
    labels = list(grouped_data.keys())
    sizes = list(grouped_data.values())
    
    # Создаем фигуру
    fig, ax = plt.subplots(figsize=(10, 8))
    
    # Цвета для секторов
    colors = plt.cm.Set3(range(len(labels)))
    
    # Создаем круговую диаграмму
    wedges, texts, autotexts = ax.pie(
        sizes, 
        labels=labels, 
        autopct='%1.1f%%',
        colors=colors,
        startangle=90,
        pctdistance=0.85
    )
    
    # Стилизация текста процентов
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_fontsize(10)
        autotext.set_weight('bold')
    
    # Добавляем центральный круг для создания "пончика"
    centre_circle = plt.Circle((0, 0), 0.70, fc='white')
    fig.gca().add_artist(centre_circle)
    
    # Заголовок
    ax.set_title(title, fontsize=14, fontweight='bold', pad=20)
    
    # Добавляем легенду с суммами
    legend_labels = [f'{l}: {v:.0f}' for l, v in zip(labels, sizes)]
    ax.legend(wedges, legend_labels, title="Расходы", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
    
    plt.tight_layout()
    
    if filename:
        plt.savefig(filename, dpi=150, bbox_inches='tight')
        print(f"Диаграмма сохранена: {filename}")
    else:
        plt.show()
    
    plt.close()


def display_summary(data):
    """Вывод сводной информации о расходах."""
    print("\n" + "=" * 60)
    print("СВОДНАЯ ИНФОРМАЦИЯ О РАСХОДАХ")
    print("=" * 60)
    
    total = sum(row[1] for row in data)
    print(f"\nОбщая сумма расходов: {total:.2f}")
    print(f"Количество записей: {len(data)}")
    
    # Группировка по категориям
    by_category = group_by_category(data)
    print("\n--- Расходы по категориям ---")
    for cat, amount in sorted(by_category.items(), key=lambda x: x[1], reverse=True):
        percent = (amount / total) * 100 if total > 0 else 0
        print(f"  {cat}: {amount:.2f} ({percent:.1f}%)")
    
    # Группировка по членам семьи
    by_member = group_by_member(data)
    print("\n--- Расходы по членам семьи ---")
    for member, amount in sorted(by_member.items(), key=lambda x: x[1], reverse=True):
        percent = (amount / total) * 100 if total > 0 else 0
        print(f"  {member}: {amount:.2f} ({percent:.1f}%)")
    
    # Группировка по датам
    by_date = group_by_date(data)
    print("\n--- Расходы по датам ---")
    for date, amount in sorted(by_date.items()):
        print(f"  {date}: {amount:.2f}")
    
    print("=" * 60)


def main():
    """Основная функция приложения."""
    print("=" * 60)
    print("ПРИЛОЖЕНИЕ ДЛЯ АНАЛИЗА РАСХОДОВ СЕМЬИ")
    print("=" * 60)
    
    # Определение пути к базе данных
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Проверяем наличие базы данных
    possible_dbs = [
        os.path.join(script_dir, "family_expenses.mdb"),
        os.path.join(script_dir, "family_expenses.accdb"),
        os.path.join(script_dir, "family_expenses.sqlite"),
        os.path.join(script_dir, "family_expenses.db"),
    ]
    
    db_path = None
    for path in possible_dbs:
        if os.path.exists(path):
            db_path = path
            break
    
    if not db_path:
        # Если база не найдена, спрашиваем путь у пользователя
        db_path = input("\nВведите путь к базе данных: ").strip().strip('"')
    
    if not os.path.exists(db_path):
        print(f"Ошибка: База данных не найдена: {db_path}")
        sys.exit(1)
    
    print(f"\nПодключение к базе данных: {db_path}")
    
    # Подключение к базе данных
    conn, db_type = connect_to_database(db_path)
    
    if conn is None:
        print("Не удалось подключиться к базе данных.")
        sys.exit(1)
    
    print(f"Тип базы данных: {db_type.upper()}")
    
    try:
        # Получение данных
        print("\nЗагрузка данных...")
        data = fetch_expenses(conn, db_type)
        
        if not data:
            print("База данных пуста!")
            sys.exit(0)
        
        # Вывод сводной информации
        display_summary(data)
        
        # Меню выбора типа группировки
        print("\nВыберите тип группировки для отображения:")
        print("  1. По категориям трат")
        print("  2. По членам семьи")
        print("  3. По датам")
        print("  4. Показать все диаграммы")
        print("  5. Выход")
        
        choice = input("\nВаш выбор (1-5): ").strip()
        
        output_dir = os.path.join(script_dir, "charts")
        os.makedirs(output_dir, exist_ok=True)
        
        if choice == '1':
            grouped = group_by_category(data)
            create_pie_chart(grouped, "Расходы по категориям", 
                           os.path.join(output_dir, "by_category.png"))
        elif choice == '2':
            grouped = group_by_member(data)
            create_pie_chart(grouped, "Расходы по членам семьи", 
                           os.path.join(output_dir, "by_member.png"))
        elif choice == '3':
            grouped = group_by_date(data)
            create_pie_chart(grouped, "Расходы по датам", 
                           os.path.join(output_dir, "by_date.png"))
        elif choice == '4':
            grouped = group_by_category(data)
            create_pie_chart(grouped, "Расходы по категориям", 
                           os.path.join(output_dir, "by_category.png"))
            grouped = group_by_member(data)
            create_pie_chart(grouped, "Расходы по членам семьи", 
                           os.path.join(output_dir, "by_member.png"))
            grouped = group_by_date(data)
            create_pie_chart(grouped, "Расходы по датам", 
                           os.path.join(output_dir, "by_date.png"))
            print(f"\nВсе диаграммы сохранены в папку: {output_dir}")
        elif choice == '5':
            print("Выход из программы.")
        else:
            print("Неверный выбор. Запустите программу снова.")
    
    finally:
        conn.close()
        print("\nСоединение с базой данных закрыто.")


if __name__ == "__main__":
    main()
