using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace FamilyExpensesApp
{
    /// <summary>
    /// Модель данных о расходах члена семьи
    /// </summary>
    public class Expense
    {
        public int Id { get; set; }
        public DateTime Date { get; set; }
        public decimal Amount { get; set; }
        public string Category { get; set; } = string.Empty;
        public string MemberName { get; set; } = string.Empty;
    }

    /// <summary>
    /// Класс для работы с базой данных MS Access
    /// </summary>
    public class DatabaseService
    {
        private readonly string _connectionString;

        public DatabaseService(string dbPath)
        {
            // Поддержка как .mdb, так и .accdb файлов
            string extension = Path.GetExtension(dbPath).ToLower();
            if (extension == ".accdb")
            {
                _connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};";
            }
            else
            {
                _connectionString = $@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={dbPath};";
            }
        }

        /// <summary>
        /// Получение всех расходов из базы данных
        /// </summary>
        public List<Expense> GetExpenses()
        {
            var expenses = new List<Expense>();

            using (var connection = new OleDbConnection(_connectionString))
            {
                connection.Open();

                string query = "SELECT Id, ExpenseDate, Amount, Category, MemberName FROM Expenses ORDER BY ExpenseDate DESC";
                
                using (var command = new OleDbCommand(query, connection))
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        expenses.Add(new Expense
                        {
                            Id = Convert.ToInt32(reader["Id"]),
                            Date = Convert.ToDateTime(reader["ExpenseDate"]),
                            Amount = Convert.ToDecimal(reader["Amount"]),
                            Category = reader["Category"].ToString() ?? string.Empty,
                            MemberName = reader["MemberName"].ToString() ?? string.Empty
                        });
                    }
                }
            }

            return expenses;
        }

        /// <summary>
        /// Группировка расходов по категориям
        /// </summary>
        public Dictionary<string, decimal> GetExpensesByCategory(List<Expense> expenses)
        {
            return expenses
                .GroupBy(e => e.Category)
                .ToDictionary(g => g.Key, g => g.Sum(e => e.Amount));
        }

        /// <summary>
        /// Группировка расходов по членам семьи
        /// </summary>
        public Dictionary<string, decimal> GetExpensesByMember(List<Expense> expenses)
        {
            return expenses
                .GroupBy(e => e.MemberName)
                .ToDictionary(g => g.Key, g => g.Sum(e => e.Amount));
        }

        /// <summary>
        /// Группировка расходов по датам
        /// </summary>
        public Dictionary<string, decimal> GetExpensesByDate(List<Expense> expenses)
        {
            return expenses
                .GroupBy(e => e.Date.Date)
                .OrderBy(g => g.Key)
                .ToDictionary(g => g.Key.ToString("dd.MM.yyyy"), g => g.Sum(e => e.Amount));
        }
    }

    /// <summary>
    /// Главная форма приложения
    /// </summary>
    public class MainForm : Form
    {
        private DataGridView _dataGridView;
        private TabControl _tabControl;
        private Button _loadButton;
        private Button _browseButton;
        private TextBox _dbPathTextBox;
        private Label _statusLabel;
        
        private List<Expense> _expenses = new();
        private DatabaseService? _dbService;

        public MainForm()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            Text = "Расходы семьи - MS Access";
            Size = new Size(1200, 800);
            StartPosition = FormStartPosition.CenterScreen;

            // Панель управления
            var controlPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 60,
                Padding = new Padding(10)
            };

            _browseButton = new Button
            {
                Text = "Обзор...",
                Location = new Point(10, 15),
                Size = new Size(80, 30)
            };
            _browseButton.Click += BrowseButton_Click;

            _dbPathTextBox = new TextBox
            {
                Location = new Point(100, 18),
                Size = new Size(600, 25),
                ReadOnly = true
            };

            _loadButton = new Button
            {
                Text = "Загрузить данные",
                Location = new Point(720, 15),
                Size = new Size(120, 30),
                Enabled = false
            };
            _loadButton.Click += LoadButton_Click;

            _statusLabel = new Label
            {
                Location = new Point(10, 35),
                Size = new Size(500, 20),
                Text = "Выберите базу данных MS Access (.mdb или .accdb)"
            };

            controlPanel.Controls.AddRange(new Control[] { 
                _browseButton, _dbPathTextBox, _loadButton, _statusLabel 
            });

            // Таблицы с графиками
            _tabControl = new TabControl
            {
                Dock = DockStyle.Fill
            };

            // Вкладка с таблицей данных
            var dataTabPage = new TabPage("Данные");
            _dataGridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };
            dataTabPage.Controls.Add(_dataGridView);

            // Вкладка с круговой диаграммой по категориям
            var categoryTabPage = CreateChartTabPage("По категориям");
            
            // Вкладка с круговой диаграммой по членам семьи
            var memberTabPage = CreateChartTabPage("По членам семьи");
            
            // Вкладка с круговой диаграммой по датам
            var dateTabPage = CreateChartTabPage("По датам");

            _tabControl.TabPages.AddRange(new TabPage[] { 
                dataTabPage, categoryTabPage, memberTabPage, dateTabPage 
            });

            Controls.Add(_tabControl);
            Controls.Add(controlPanel);
        }

        private TabPage CreateChartTabPage(string title)
        {
            var tabPage = new TabPage(title);
            
            var chart = new Chart
            {
                Name = $"chart_{title.Replace(" ", "_")}",
                Dock = DockStyle.Fill
            };

            chart.ChartAreas.Add(new ChartArea());
            chart.Legends.Add(new Legend());
            chart.Titles.Add(new Title(title));

            // Настройка круговой диаграммы
            var series = new Series
            {
                ChartType = SeriesChartType.Pie,
                IsValueShownAsLabel = true,
                Label = "#PERCENT{P1}\n#VAL руб.",
                LegendText = "#AXISLABEL"
            };
            chart.Series.Add(series);

            // Цвета для диаграммы
            Color[] colors = {
                Color.FromArgb(255, 99, 132),
                Color.FromArgb(54, 162, 235),
                Color.FromArgb(255, 206, 86),
                Color.FromArgb(75, 192, 192),
                Color.FromArgb(153, 102, 255),
                Color.FromArgb(255, 159, 64),
                Color.FromArgb(199, 199, 199),
                Color.FromArgb(83, 102, 255),
                Color.FromArgb(255, 99, 255),
                Color.FromArgb(99, 255, 132)
            };

            foreach (var color in colors)
            {
                series.Points.Add(new DataPoint { Color = color });
            }

            tabPage.Controls.Add(chart);
            tabPage.Tag = chart; // Сохраняем ссылку на график

            return tabPage;
        }

        private void BrowseButton_Click(object? sender, EventArgs e)
        {
            using var openFileDialog = new OpenFileDialog
            {
                Filter = "Базы данных MS Access (*.mdb;*.accdb)|*.mdb;*.accdb|Все файлы (*.*)|*.*",
                Title = "Выберите базу данных MS Access"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                _dbPathTextBox.Text = openFileDialog.FileName;
                _loadButton.Enabled = true;
                _statusLabel.Text = $"Файл выбран: {Path.GetFileName(openFileDialog.FileName)}";
            }
        }

        private void LoadButton_Click(object? sender, EventArgs e)
        {
            try
            {
                string dbPath = _dbPathTextBox.Text;
                if (!File.Exists(dbPath))
                {
                    MessageBox.Show("Файл базы данных не найден!", "Ошибка", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                _dbService = new DatabaseService(dbPath);
                _expenses = _dbService.GetExpenses();

                DisplayData();
                UpdateCharts();

                _statusLabel.Text = $"Загружено {_expenses.Count} записей. Общая сумма: {_expenses.Sum(x => x.Amount):F2} руб.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных:\n{ex.Message}", "Ошибка", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DisplayData()
        {
            _dataGridView.Columns.Clear();
            _dataGridView.Columns.Add("Date", "Дата");
            _dataGridView.Columns.Add("MemberName", "Член семьи");
            _dataGridView.Columns.Add("Category", "Категория");
            _dataGridView.Columns.Add("Amount", "Сумма (руб.)");

            _dataGridView.Rows.Clear();

            foreach (var expense in _expenses)
            {
                _dataGridView.Rows.Add(
                    expense.Date.ToString("dd.MM.yyyy"),
                    expense.MemberName,
                    expense.Category,
                    expense.Amount.ToString("F2")
                );
            }
        }

        private void UpdateCharts()
        {
            // Обновление графика по категориям
            UpdateChart("По категориям", _dbService!.GetExpensesByCategory(_expenses));
            
            // Обновление графика по членам семьи
            UpdateChart("По членам семьи", _dbService.GetExpensesByMember(_expenses));
            
            // Обновление графика по датам
            UpdateChart("По датам", _dbService.GetExpensesByDate(_expenses));
        }

        private void UpdateChart(string tabTitle, Dictionary<string, decimal> data)
        {
            var tabPage = _tabControl.TabPages.Cast<TabPage>()
                .FirstOrDefault(t => t.Text == tabTitle);
            
            if (tabPage?.Tag is not Chart chart) return;

            chart.Series[0].Points.Clear();
            chart.Titles[0].Text = $"{tabTitle} (Всего: {data.Values.Sum():F2} руб.)";

            int pointIndex = 0;
            Color[] colors = {
                Color.FromArgb(255, 99, 132),
                Color.FromArgb(54, 162, 235),
                Color.FromArgb(255, 206, 86),
                Color.FromArgb(75, 192, 192),
                Color.FromArgb(153, 102, 255),
                Color.FromArgb(255, 159, 64),
                Color.FromArgb(199, 199, 199),
                Color.FromArgb(83, 102, 255),
                Color.FromArgb(255, 99, 255),
                Color.FromArgb(99, 255, 132),
                Color.FromArgb(255, 193, 7),
                Color.FromArgb(76, 175, 80),
                Color.FromArgb(33, 150, 243),
                Color.FromArgb(156, 39, 176),
                Color.FromArgb(233, 30, 99)
            };

            foreach (var item in data)
            {
                var point = new DataPoint
                {
                    AxisLabel = item.Key,
                    YValues = new[] { (double)item.Value },
                    Color = colors[pointIndex % colors.Length],
                    Label = "#PERCENT{P1}\n#VAL руб."
                };
                chart.Series[0].Points.Add(point);
                pointIndex++;
            }

            // Автоподписи легенды
            chart.Legends[0].Enabled = true;
            chart.ChartAreas[0].Area3DStyle.Enable3D = false;
        }

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault();
            Application.Run(new MainForm());
        }
    }
}
