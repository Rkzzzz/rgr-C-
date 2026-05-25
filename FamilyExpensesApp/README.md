# Приложение для работы с расходами семьи (C# + MS Access)

## Описание

Приложение на C# для чтения данных о расходах членов семьи из базы данных MS Access и отображения их в виде круговых диаграмм с группировкой по:
- Дате
- Сумме
- Категории трат (еда, транспорт, развлечения и т.п.)
- Имени члена семьи

## Структура проекта

```
FamilyExpensesApp/
├── FamilyExpensesApp.csproj    # Файл проекта .NET
├── Program.cs                  # Основной код приложения
├── create_test_db.py          # Скрипт для создания тестовой БД
└── README.md                  # Этот файл
```

## Требования

### Для запуска приложения:
- **.NET 8.0 SDK** (или выше)
- **Windows** (требуется для работы с Windows Forms и MS Access)
- **Microsoft Access Database Engine** (для подключения к .mdb/.accdb файлам)
  - [Microsoft Access Database Engine 2016 Redistributable](https://www.microsoft.com/en-us/download/details.aspx?id=54920)

### Структура таблицы в БД MS Access

Таблица должна называться **Expenses** и содержать следующие поля:

| Поле | Тип данных | Описание |
|------|-----------|----------|
| Id | Autonumber (или Integer) | Первичный ключ |
| ExpenseDate | Date/Time | Дата расхода |
| Amount | Currency (или Number) | Сумма расхода |
| Category | Text | Категория (Еда, Транспорт, Развлечения и т.д.) |
| MemberName | Text | Имя члена семьи |

## Установка и запуск

### 1. Создание тестовой базы данных (опционально)

Для тестирования можно создать SQLite базу данных:

```bash
python create_test_db.py
```

> **Примечание:** Приложение предназначено для работы с MS Access (.mdb, .accdb), но для тестирования логики можно использовать SQLite.

### 2. Сборка проекта

```bash
cd FamilyExpensesApp
dotnet restore
dotnet build
```

### 3. Запуск приложения

```bash
dotnet run
```

Или создайте исполняемый файл:

```bash
dotnet publish -c Release -r win-x64 --self-contained false
```

## Использование

1. Запустите приложение
2. Нажмите кнопку **"Обзор..."** для выбора файла базы данных MS Access (.mdb или .accdb)
3. Нажмите **"Загрузить данные"**
4. Просматривайте данные на вкладках:
   - **Данные** - таблица со всеми расходами
   - **По категориям** - круговая диаграмма расходов по категориям
   - **По членам семьи** - круговая диаграмма расходов по каждому члену семьи
   - **По датам** - круговая диаграмма расходов по датам

## Особенности

- Поддержка форматов **.mdb** (Jet Engine) и **.accdb** (ACE Engine)
- Автоматическое вычисление процентов на диаграммах
- Цветовое кодирование секторов диаграммы
- Отображение общей суммы расходов
- Адаптивный интерфейс с вкладками

## Пример SQL-запроса для создания таблицы в MS Access

```sql
CREATE TABLE Expenses (
    Id AUTOINCREMENT PRIMARY KEY,
    ExpenseDate DATETIME NOT NULL,
    Amount CURRENCY NOT NULL,
    Category TEXT(50) NOT NULL,
    MemberName TEXT(100) NOT NULL
);
```

## Зависимости

- `System.Data.OleDb` - для подключения к MS Access
- `System.Windows.Forms.DataVisualization` - для построения диаграмм
- `System.Windows.Forms` - для GUI приложения

## Примечания

- Приложение работает только на **Windows** (требуется для Windows Forms и OleDb провайдера)
- Для работы с .accdb файлами необходим установленный Microsoft Access Database Engine
- На Linux/Mac возможно использование только через Mono с ограничениями
