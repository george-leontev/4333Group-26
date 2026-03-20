using System.IO;
using System.Windows;
using Microsoft.Data.SqlClient;
using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

// Явно указываем какой LicenseContext использовать — убирает ошибку "ambiguous reference"
using ExcelLicenseContext = OfficeOpenXml.LicenseContext;

namespace Group4333
{
    public partial class Group4333_Leontev : Window
    {
        // ⚠️ Замени на имя своего сервера из SSMS
        // Примеры: ".\SQLEXPRESS"  /  "localhost"  /  "DESKTOP-XXX\SQLEXPRESS"
        private readonly string _connectionString =
            "Server=localhost;Database=ISRPOOrdersDB;Trusted_Connection=True;TrustServerCertificate=True;";

        private List<Order> _orders = new();

        public Group4333_Leontev()
        {
            InitializeComponent();
            // Используем псевдоним чтобы избежать конфликта с System.ComponentModel.LicenseContext
            ExcelPackage.LicenseContext = ExcelLicenseContext.NonCommercial;
            InitializeDatabase();
        }

        // ──────────────────────────────────────────────
        //  БАЗА ДАННЫХ — SQL SERVER
        // ──────────────────────────────────────────────

        private void InitializeDatabase()
        {
            try
            {
                using var conn = new SqlConnection(_connectionString);
                conn.Open();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = @"
                    IF NOT EXISTS (
                        SELECT * FROM sysobjects
                        WHERE name = 'Orders' AND xtype = 'U'
                    )
                    CREATE TABLE Orders (
                        Id          INT            PRIMARY KEY,
                        OrderCode   NVARCHAR(50)   NOT NULL,
                        CreatedDate DATE           NOT NULL,
                        ClientCode  NVARCHAR(50)   NOT NULL,
                        Services    NVARCHAR(200)  NOT NULL,
                        Status      NVARCHAR(50)   NOT NULL
                    );";
                cmd.ExecuteNonQuery();
                DbStatusText.Text = "База данных: SQL Server — OrdersDB ✅";
            }
            catch (Exception ex)
            {
                DbStatusText.Text = "❌ Нет подключения к SQL Server";
                MessageBox.Show(
                    $"Ошибка подключения к SQL Server:\n{ex.Message}\n\n" +
                    $"Проверь строку подключения:\n{_connectionString}",
                    "Ошибка БД", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveOrdersToDatabase(List<Order> orders)
        {
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var transaction = conn.BeginTransaction();
            try
            {
                using (var del = conn.CreateCommand())
                {
                    del.Transaction = transaction;
                    del.CommandText = "DELETE FROM Orders;";
                    del.ExecuteNonQuery();
                }

                foreach (var o in orders)
                {
                    using var cmd = conn.CreateCommand();
                    cmd.Transaction = transaction;
                    cmd.CommandText = @"
                        INSERT INTO Orders (Id, OrderCode, CreatedDate, ClientCode, Services, Status)
                        VALUES (@id, @code, @date, @client, @services, @status);";
                    cmd.Parameters.AddWithValue("@id", o.Id);
                    cmd.Parameters.AddWithValue("@code", o.OrderCode);
                    cmd.Parameters.AddWithValue("@date", o.CreatedDate);
                    cmd.Parameters.AddWithValue("@client", o.ClientCode);
                    cmd.Parameters.AddWithValue("@services", o.Services);
                    cmd.Parameters.AddWithValue("@status", o.Status);
                    cmd.ExecuteNonQuery();
                }

                transaction.Commit();
            }
            catch
            {
                transaction.Rollback();
                throw;
            }
        }

        // ──────────────────────────────────────────────
        //  ИМПОРТ ИЗ EXCEL
        // ──────────────────────────────────────────────

        private void ImportData_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title = "Выберите файл 2.xlsx",
                Filter = "Excel файлы (*.xlsx)|*.xlsx",
                FileName = "2.xlsx"
            };

            if (dlg.ShowDialog() != true) return;

            try
            {
                StatusText.Text = "Загрузка...";

                var orders = ReadOrdersFromExcel(dlg.FileName);
                SaveOrdersToDatabase(orders);

                _orders = orders;
                OrdersGrid.ItemsSource = _orders;

                CountText.Text =
                    $"Загружено: {_orders.Count}  |  " +
                    $"Активных: {_orders.Count(o => o.Status == "Активен")}  |  " +
                    $"Завершённых: {_orders.Count(o => o.Status == "Завершён")}  |  " +
                    $"Отменённых: {_orders.Count(o => o.Status == "Отменён")}";

                StatusText.Text =
                    $"✅ Импортировано {orders.Count} записей из {Path.GetFileName(dlg.FileName)}";
            }
            catch (Exception ex)
            {
                StatusText.Text = "❌ Ошибка импорта";
                MessageBox.Show($"Ошибка при чтении файла:\n{ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private List<Order> ReadOrdersFromExcel(string filePath)
        {
            var result = new List<Order>();

            using var package = new ExcelPackage(new FileInfo(filePath));
            var ws = package.Workbook.Worksheets[0];
            int rowCount = ws.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++)
            {
                var idVal = ws.Cells[row, 1].Value;
                var code = ws.Cells[row, 2].Value?.ToString()?.Trim();
                var dateVal = ws.Cells[row, 3].Value;
                var client = ws.Cells[row, 4].Value?.ToString()?.Trim();
                var service = ws.Cells[row, 5].Value?.ToString()?.Trim();
                var status = ws.Cells[row, 6].Value?.ToString()?.Trim();

                if (idVal == null || string.IsNullOrEmpty(code)) continue;

                DateTime createdDate;
                if (dateVal is double d)
                    createdDate = DateTime.FromOADate(d);
                else if (!DateTime.TryParse(dateVal?.ToString(), out createdDate))
                    createdDate = DateTime.Today;

                result.Add(new Order
                {
                    Id = Convert.ToInt32(idVal),
                    OrderCode = code ?? "",
                    CreatedDate = createdDate,
                    ClientCode = client ?? "",
                    Services = service ?? "",
                    Status = status ?? ""
                });
            }

            return result;
        }

        // ──────────────────────────────────────────────
        //  ЭКСПОРТ В EXCEL (группировка по Статусу)
        // ──────────────────────────────────────────────

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            if (_orders.Count == 0)
            {
                MessageBox.Show("Сначала импортируйте данные!", "Внимание",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dlg = new SaveFileDialog
            {
                Title = "Сохранить экспорт",
                Filter = "Excel файлы (*.xlsx)|*.xlsx",
                FileName = $"export_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"
            };

            if (dlg.ShowDialog() != true) return;

            try
            {
                StatusText.Text = "Экспорт...";
                ExportGroupedByStatus(dlg.FileName);
                StatusText.Text = $"✅ Экспортировано: {Path.GetFileName(dlg.FileName)}";

                var res = MessageBox.Show("Экспорт завершён! Открыть файл?", "Успех",
                    MessageBoxButton.YesNo, MessageBoxImage.Information);

                if (res == MessageBoxResult.Yes)
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = dlg.FileName,
                        UseShellExecute = true
                    });
            }
            catch (Exception ex)
            {
                StatusText.Text = "❌ Ошибка экспорта";
                MessageBox.Show($"Ошибка при экспорте:\n{ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportGroupedByStatus(string filePath)
        {
            using var package = new ExcelPackage();

            var statusRowColor = new Dictionary<string, Color>
            {
                { "Активен",  Color.FromArgb(0xD9, 0xEA, 0xD3) },
                { "Завершён", Color.FromArgb(0xCF, 0xE2, 0xF3) },
                { "Отменён",  Color.FromArgb(0xF4, 0xCC, 0xCC) }
            };

            var statusHeaderColor = new Dictionary<string, Color>
            {
                { "Активен",  Color.FromArgb(0x38, 0x76, 0x1D) },
                { "Завершён", Color.FromArgb(0x21, 0x66, 0xA6) },
                { "Отменён",  Color.FromArgb(0xCC, 0x00, 0x00) }
            };

            string[] columns = { "Id", "Код заказа", "Дата создания", "Код клиента", "Услуги", "Статус" };

            foreach (var group in _orders.GroupBy(o => o.Status).OrderBy(g => g.Key))
            {
                var ws = package.Workbook.Worksheets.Add(group.Key);

                Color rowBg = statusRowColor.GetValueOrDefault(group.Key, Color.LightGray);
                Color headerBg = statusHeaderColor.GetValueOrDefault(group.Key, Color.DarkGray);

                // Строка 1 — заголовок листа
                ws.Cells[1, 1, 1, 6].Merge = true;
                ws.Cells[1, 1].Value = $"Заказы — статус: {group.Key}  ({group.Count()} записей)";
                ws.Cells[1, 1].Style.Font.Size = 13;
                ws.Cells[1, 1].Style.Font.Bold = true;
                ws.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(headerBg);
                ws.Cells[1, 1].Style.Font.Color.SetColor(Color.White);
                ws.Row(1).Height = 26;

                // Строка 2 — заголовки столбцов
                for (int col = 0; col < columns.Length; col++)
                {
                    var cell = ws.Cells[2, col + 1];
                    cell.Value = columns[col];
                    cell.Style.Font.Bold = true;
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x44, 0x72, 0xC4));
                    cell.Style.Font.Color.SetColor(Color.White);
                    ApplyBorder(cell);
                }
                ws.Row(2).Height = 22;

                // Строки данных
                int dataRow = 3;
                foreach (var o in group.OrderBy(x => x.Id))
                {
                    ws.Cells[dataRow, 1].Value = o.Id;
                    ws.Cells[dataRow, 2].Value = o.OrderCode;
                    ws.Cells[dataRow, 3].Value = o.CreatedDate.ToString("dd.MM.yyyy");
                    ws.Cells[dataRow, 4].Value = o.ClientCode;
                    ws.Cells[dataRow, 5].Value = o.Services;
                    ws.Cells[dataRow, 6].Value = o.Status;

                    Color bg = dataRow % 2 == 0 ? rowBg : Color.White;
                    for (int col = 1; col <= 6; col++)
                    {
                        var cell = ws.Cells[dataRow, col];
                        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(bg);
                        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ApplyBorder(cell);
                    }
                    ws.Row(dataRow).Height = 20;
                    dataRow++;
                }

                // Итоговая строка
                ws.Cells[dataRow, 1, dataRow, 5].Merge = true;
                ws.Cells[dataRow, 1].Value = "Итого записей:";
                ws.Cells[dataRow, 1].Style.Font.Bold = true;
                ws.Cells[dataRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells[dataRow, 6].Value = group.Count();
                ws.Cells[dataRow, 6].Style.Font.Bold = true;
                ws.Cells[dataRow, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                for (int col = 1; col <= 6; col++)
                {
                    ws.Cells[dataRow, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[dataRow, col].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xFF, 0xFF, 0xCC));
                    ApplyBorder(ws.Cells[dataRow, col]);
                }

                // Ширина колонок
                ws.Column(1).Width = 8;
                ws.Column(2).Width = 16;
                ws.Column(3).Width = 18;
                ws.Column(4).Width = 16;
                ws.Column(5).Width = 24;
                ws.Column(6).Width = 14;
            }

            package.SaveAs(new FileInfo(filePath));
        }

        private static void ApplyBorder(ExcelRange cell)
        {
            cell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            cell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
        }
    }

    public class Order
    {
        public int Id { get; set; }
        public string OrderCode { get; set; } = "";
        public DateTime CreatedDate { get; set; }
        public string ClientCode { get; set; } = "";
        public string Services { get; set; } = "";
        public string Status { get; set; } = "";
    }
}
