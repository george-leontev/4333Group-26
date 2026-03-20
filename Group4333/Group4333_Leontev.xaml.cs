// ── Псевдонимы для устранения конфликтов имён ──────────────────────────────
using DrawingColor = System.Drawing.Color;
using WordColor = DocumentFormat.OpenXml.Wordprocessing.Color;
using WordFontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;
using ExcelFontSize = OfficeOpenXml.Style.ExcelFont;
using ExcelLicenseContext = OfficeOpenXml.LicenseContext;
// ────────────────────────────────────────────────────────────────────────────

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Data.SqlClient;
using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Group4333
{
    public partial class Group4333_Leontev : Window
    {
        private readonly string _connectionString =
            "Server=localhost;Database=ISRPOOrdersDB;Trusted_Connection=True;TrustServerCertificate=True;";

        private List<Order> _orders = new();

        public Group4333_Leontev()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = ExcelLicenseContext.NonCommercial;
            InitializeDatabase();
        }

        private void InitializeDatabase()
        {
            try
            {
                using var conn = new SqlConnection(_connectionString);
                conn.Open();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = @"
                    IF NOT EXISTS (
                        SELECT * FROM sysobjects WHERE name='Orders' AND xtype='U'
                    )
                    CREATE TABLE Orders (
                        Id          INT            PRIMARY KEY,
                        OrderCode   NVARCHAR(50)   NOT NULL,
                        CreatedDate DATE           NOT NULL,
                        ClientCode  NVARCHAR(50)   NOT NULL,
                        Services    NVARCHAR(200)  NOT NULL,
                        Status      NVARCHAR(50)   NOT NULL,
                        Source      NVARCHAR(10)   NOT NULL DEFAULT 'xlsx'
                    );
                    IF NOT EXISTS (
                        SELECT * FROM sys.columns
                        WHERE object_id = OBJECT_ID('Orders') AND name = 'Source'
                    )
                    ALTER TABLE Orders ADD Source NVARCHAR(10) NOT NULL DEFAULT 'xlsx';";
                cmd.ExecuteNonQuery();
                DbStatusText.Text = "База данных: SQL Server — OrdersDB ✅";
            }
            catch (Exception ex)
            {
                DbStatusText.Text = "❌ Нет подключения к SQL Server";
                MessageBox.Show(
                    $"Ошибка подключения к SQL Server:\n{ex.Message}\n\n" +
                    $"Строка подключения:\n{_connectionString}",
                    "Ошибка БД", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveOrdersToDatabase(List<Order> orders, string source)
        {
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var transaction = conn.BeginTransaction();
            try
            {
                using (var del = conn.CreateCommand())
                {
                    del.Transaction = transaction;
                    del.CommandText = "DELETE FROM Orders WHERE Source = @source;";
                    del.Parameters.AddWithValue("@source", source);
                    del.ExecuteNonQuery();
                }

                foreach (var o in orders)
                {
                    using var cmd = conn.CreateCommand();
                    cmd.Transaction = transaction;
                    cmd.CommandText = @"
                        IF EXISTS (SELECT 1 FROM Orders WHERE Id = @id)
                            UPDATE Orders
                            SET OrderCode=@code, CreatedDate=@date, ClientCode=@client,
                                Services=@services, Status=@status, Source=@source
                            WHERE Id = @id
                        ELSE
                            INSERT INTO Orders
                                (Id, OrderCode, CreatedDate, ClientCode, Services, Status, Source)
                            VALUES
                                (@id, @code, @date, @client, @services, @status, @source);";
                    cmd.Parameters.AddWithValue("@id", o.Id);
                    cmd.Parameters.AddWithValue("@code", o.OrderCode);
                    cmd.Parameters.AddWithValue("@date", o.CreatedDate);
                    cmd.Parameters.AddWithValue("@client", o.ClientCode);
                    cmd.Parameters.AddWithValue("@services", o.Services);
                    cmd.Parameters.AddWithValue("@status", o.Status);
                    cmd.Parameters.AddWithValue("@source", source);
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

        private void RefreshGrid()
        {
            CountText.Text =
                $"Всего: {_orders.Count}  |  " +
                $"Активных: {_orders.Count(o => o.Status == "Активен")}  |  " +
                $"Завершённых: {_orders.Count(o => o.Status == "Завершён")}  |  " +
                $"Отменённых: {_orders.Count(o => o.Status == "Отменён")}";
            OrdersGrid.ItemsSource = null;
            OrdersGrid.ItemsSource = _orders;
        }
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
                StatusText.Text = "Загрузка xlsx...";
                var orders = ReadOrdersFromExcel(dlg.FileName);
                SaveOrdersToDatabase(orders, "xlsx");
                _orders = _orders.Where(o => o.Source == "json")
                                 .Concat(orders)
                                 .OrderBy(o => o.Id).ToList();
                RefreshGrid();
                StatusText.Text = $"✅ Excel: импортировано {orders.Count} записей";
            }
            catch (Exception ex)
            {
                StatusText.Text = "❌ Ошибка импорта xlsx";
                MessageBox.Show($"Ошибка:\n{ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
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

                DateTime dt;
                if (dateVal is double d) dt = DateTime.FromOADate(d);
                else if (!DateTime.TryParse(dateVal?.ToString(), out dt)) dt = DateTime.Today;

                result.Add(new Order
                {
                    Id = Convert.ToInt32(idVal),
                    OrderCode = code ?? "",
                    CreatedDate = dt,
                    ClientCode = client ?? "",
                    Services = service ?? "",
                    Status = status ?? "",
                    Source = "xlsx"
                });
            }
            return result;
        }
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
                Title = "Сохранить экспорт Excel",
                Filter = "Excel файлы (*.xlsx)|*.xlsx",
                FileName = $"export_excel_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"
            };
            if (dlg.ShowDialog() != true) return;
            try
            {
                StatusText.Text = "Экспорт в Excel...";
                ExportGroupedByStatusExcel(dlg.FileName);
                StatusText.Text = $"✅ Excel: {Path.GetFileName(dlg.FileName)}";
                AskOpen(dlg.FileName);
            }
            catch (Exception ex)
            {
                StatusText.Text = "❌ Ошибка экспорта Excel";
                MessageBox.Show($"Ошибка:\n{ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportGroupedByStatusExcel(string filePath)
        {
            using var package = new ExcelPackage();

            var rowColors = new Dictionary<string, DrawingColor>
            {
                { "Активен",  DrawingColor.FromArgb(0xD9, 0xEA, 0xD3) },
                { "Завершён", DrawingColor.FromArgb(0xCF, 0xE2, 0xF3) },
                { "Отменён",  DrawingColor.FromArgb(0xF4, 0xCC, 0xCC) }
            };
            var hdrColors = new Dictionary<string, DrawingColor>
            {
                { "Активен",  DrawingColor.FromArgb(0x38, 0x76, 0x1D) },
                { "Завершён", DrawingColor.FromArgb(0x21, 0x66, 0xA6) },
                { "Отменён",  DrawingColor.FromArgb(0xCC, 0x00, 0x00) }
            };

            string[] cols = { "Id", "Код заказа", "Дата создания",
                               "Код клиента", "Услуги", "Статус" };

            foreach (var group in _orders.GroupBy(o => o.Status).OrderBy(g => g.Key))
            {
                var ws = package.Workbook.Worksheets.Add(group.Key);
                DrawingColor rowBg = rowColors.GetValueOrDefault(group.Key, DrawingColor.LightGray);
                DrawingColor hdrBg = hdrColors.GetValueOrDefault(group.Key, DrawingColor.DarkGray);

                ws.Cells[1, 1, 1, 6].Merge = true;
                ws.Cells[1, 1].Value = $"Заказы — статус: {group.Key}  ({group.Count()} записей)";
                ws.Cells[1, 1].Style.Font.Size = 13;
                ws.Cells[1, 1].Style.Font.Bold = true;
                ws.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(hdrBg);
                ws.Cells[1, 1].Style.Font.Color.SetColor(DrawingColor.White);
                ws.Row(1).Height = 26;

                for (int c = 0; c < cols.Length; c++)
                {
                    var cell = ws.Cells[2, c + 1];
                    cell.Value = cols[c];
                    cell.Style.Font.Bold = true;
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(DrawingColor.FromArgb(0x44, 0x72, 0xC4));
                    cell.Style.Font.Color.SetColor(DrawingColor.White);
                    ApplyExcelBorder(cell);
                }
                ws.Row(2).Height = 22;

                int dataRow = 3;
                foreach (var o in group.OrderBy(x => x.Id))
                {
                    ws.Cells[dataRow, 1].Value = o.Id;
                    ws.Cells[dataRow, 2].Value = o.OrderCode;
                    ws.Cells[dataRow, 3].Value = o.CreatedDate.ToString("dd.MM.yyyy");
                    ws.Cells[dataRow, 4].Value = o.ClientCode;
                    ws.Cells[dataRow, 5].Value = o.Services;
                    ws.Cells[dataRow, 6].Value = o.Status;

                    DrawingColor bg = dataRow % 2 == 0 ? rowBg : DrawingColor.White;
                    for (int c = 1; c <= 6; c++)
                    {
                        var cell = ws.Cells[dataRow, c];
                        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(bg);
                        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ApplyExcelBorder(cell);
                    }
                    dataRow++;
                }
                
                ws.Cells[dataRow, 1, dataRow, 5].Merge = true;
                ws.Cells[dataRow, 1].Value = "Итого:";
                ws.Cells[dataRow, 1].Style.Font.Bold = true;
                ws.Cells[dataRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells[dataRow, 6].Value = group.Count();
                ws.Cells[dataRow, 6].Style.Font.Bold = true;
                ws.Cells[dataRow, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                for (int c = 1; c <= 6; c++)
                {
                    ws.Cells[dataRow, c].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[dataRow, c].Style.Fill.BackgroundColor.SetColor(
                        DrawingColor.FromArgb(0xFF, 0xFF, 0xCC));
                    ApplyExcelBorder(ws.Cells[dataRow, c]);
                }

                ws.Column(1).Width = 8; ws.Column(2).Width = 16;
                ws.Column(3).Width = 18; ws.Column(4).Width = 16;
                ws.Column(5).Width = 24; ws.Column(6).Width = 14;
            }
            package.SaveAs(new FileInfo(filePath));
        }

        private static void ApplyExcelBorder(OfficeOpenXml.ExcelRange cell)
        {
            cell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            cell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
        }
        private void ImportJson_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title = "Выберите файл 2.json",
                Filter = "JSON файлы (*.json)|*.json",
                FileName = "2.json"
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                StatusText.Text = "Загрузка JSON...";
                var orders = ReadOrdersFromJson(dlg.FileName);
                SaveOrdersToDatabase(orders, "json");
                _orders = _orders.Where(o => o.Source == "xlsx")
                                 .Concat(orders)
                                 .OrderBy(o => o.Id).ToList();
                RefreshGrid();
                StatusText.Text = $"✅ JSON: импортировано {orders.Count} записей";
            }
            catch (Exception ex)
            {
                StatusText.Text = "❌ Ошибка импорта JSON";
                MessageBox.Show($"Ошибка:\n{ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private List<Order> ReadOrdersFromJson(string filePath)
        {
            var json = File.ReadAllText(filePath, System.Text.Encoding.UTF8);
            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            var raw = JsonSerializer.Deserialize<List<JsonOrder>>(json, options)
                          ?? throw new Exception("Не удалось десериализовать JSON");

            return raw.Select(r => new Order
            {
                Id = r.Id,
                OrderCode = r.OrderCode ?? "",
                CreatedDate = DateTime.TryParse(r.CreatedDate, out var dt) ? dt : DateTime.Today,
                ClientCode = r.ClientCode ?? "",
                Services = r.Services ?? "",
                Status = r.Status ?? "",
                Source = "json"
            }).ToList();
        }
        private void ExportToWord_Click(object sender, RoutedEventArgs e)
        {
            if (_orders.Count == 0)
            {
                MessageBox.Show("Сначала импортируйте данные!", "Внимание",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            var dlg = new SaveFileDialog
            {
                Title = "Сохранить экспорт Word",
                Filter = "Word документы (*.docx)|*.docx",
                FileName = $"export_word_{DateTime.Now:yyyyMMdd_HHmm}.docx"
            };
            if (dlg.ShowDialog() != true) return;
            try
            {
                StatusText.Text = "Экспорт в Word...";
                ExportGroupedByStatusWord(dlg.FileName);
                StatusText.Text = $"✅ Word: {Path.GetFileName(dlg.FileName)}";
                AskOpen(dlg.FileName);
            }
            catch (Exception ex)
            {
                StatusText.Text = "❌ Ошибка экспорта Word";
                MessageBox.Show($"Ошибка:\n{ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void ExportGroupedByStatusWord(string filePath)
        {
            using var wordDoc = WordprocessingDocument.Create(
                filePath, WordprocessingDocumentType.Document);

            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;

            AddWordStyles(mainPart);

            // Цвета статусов (hex-строки для Word XML)
            var accentColors = new Dictionary<string, string>
            {
                { "Активен",  "38761D" },
                { "Завершён", "2166A6" },
                { "Отменён",  "CC0000" }
            };

            var groups = _orders.GroupBy(o => o.Status).OrderBy(g => g.Key).ToList();

            for (int gi = 0; gi < groups.Count; gi++)
            {
                var group = groups[gi];
                string hex = accentColors.GetValueOrDefault(group.Key, "4472C4");

                body.AppendChild(MakeWordHeading(
                    $"Заказы — статус: {group.Key}", hex));

                body.AppendChild(MakeWordParagraph(
                    $"Количество записей: {group.Count()}    |    " +
                    $"Дата отчёта: {DateTime.Now:dd.MM.yyyy HH:mm}",
                    "666666", 18));

                body.AppendChild(new Paragraph());

                body.AppendChild(MakeWordTable(
                    group.OrderBy(x => x.Id).ToList(), hex));

                if (gi < groups.Count - 1)
                    body.AppendChild(MakeWordPageBreak());
            }

            mainPart.Document.Save();
        }

        private static Paragraph MakeWordHeading(string text, string colorHex)
        {
            var pPr = new ParagraphProperties(
                new ParagraphBorders(                  
                    new BottomBorder
                    {
                        Val = BorderValues.Single,
                        Size = 8,
                        Color = colorHex,
                        Space = 4
                    }
                ),
                new SpacingBetweenLines { After = "120" }
            );

            var rPr = new RunProperties(
                new Bold(),
                new WordFontSize { Val = "36" },        
                new WordColor { Val = colorHex },         
                new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }
            );

            return new Paragraph(pPr, new Run(rPr, new Text(text)));
        }

        private static Paragraph MakeWordParagraph(string text, string colorHex,
            int halfPt = 20, bool bold = false)
        {
            var rPr = new RunProperties(
                new WordFontSize { Val = halfPt.ToString() },
                new WordColor { Val = colorHex },
                new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }
            );
            if (bold) rPr.AppendChild(new Bold());

            return new Paragraph(
                new ParagraphProperties(new SpacingBetweenLines { After = "60" }),
                new Run(rPr, new Text(text))
            );
        }

        private static Table MakeWordTable(List<Order> orders, string accentHex)
        {
            int[] widths = { 700, 1400, 1500, 1400, 2560, 1800 };
            string[] headers = { "Id", "Код заказа", "Дата создания",
                                  "Код клиента", "Услуги", "Статус" };

            var table = new Table();
            table.AppendChild(new TableProperties(
                new TableWidth { Width = "9360", Type = TableWidthUnitValues.Dxa },
                new TableBorders(
                    MakeWordBorder<TopBorder>(),
                    MakeWordBorder<BottomBorder>(),
                    MakeWordBorder<LeftBorder>(),
                    MakeWordBorder<RightBorder>(),
                    MakeWordBorder<InsideHorizontalBorder>(),
                    MakeWordBorder<InsideVerticalBorder>()
                )
            ));

            var hdrRow = new TableRow();
            foreach (var (h, w) in headers.Zip(widths))
                hdrRow.AppendChild(MakeWordCell(h, w, accentHex, "FFFFFF", true));
            table.AppendChild(hdrRow);

            bool alt = false;
            foreach (var o in orders)
            {
                string bg = alt ? "F2F2F2" : "FFFFFF";
                var row = new TableRow();
                row.AppendChild(MakeWordCell(o.Id.ToString(), widths[0], bg, "000000", false));
                row.AppendChild(MakeWordCell(o.OrderCode, widths[1], bg, "000000", false));
                row.AppendChild(MakeWordCell(o.CreatedDate.ToString("dd.MM.yyyy"), widths[2], bg, "000000", false));
                row.AppendChild(MakeWordCell(o.ClientCode, widths[3], bg, "000000", false));
                row.AppendChild(MakeWordCell(o.Services, widths[4], bg, "000000", false));
                row.AppendChild(MakeWordCell(o.Status, widths[5], bg, "000000", false));
                table.AppendChild(row);
                alt = !alt;
            }

            int totalW = widths[0] + widths[1] + widths[2] + widths[3] + widths[4];
            var totalRow = new TableRow();
            totalRow.AppendChild(MakeWordCell("Итого:", totalW, "FFFFCC", "000000", true, 5));
            totalRow.AppendChild(MakeWordCell(orders.Count.ToString(), widths[5], "FFFFCC", "000000", true));
            table.AppendChild(totalRow);

            return table;
        }

        private static TableCell MakeWordCell(string text, int widthDxa,
            string fillHex, string fontColorHex, bool bold, int colSpan = 1)
        {
            var cellProps = new TableCellProperties(
                new TableCellWidth { Width = widthDxa.ToString(), Type = TableWidthUnitValues.Dxa },
                new Shading { Fill = fillHex, Val = ShadingPatternValues.Clear },
                new TableCellMargin(
                    new TopMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                    new BottomMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                    new LeftMargin { Width = "120", Type = TableWidthUnitValues.Dxa },
                    new RightMargin { Width = "120", Type = TableWidthUnitValues.Dxa }
                )
            );
            if (colSpan > 1)
                cellProps.AppendChild(new GridSpan { Val = colSpan });

            var rPr = new RunProperties(
                new WordFontSize { Val = "20" },
                new WordColor { Val = fontColorHex },
                new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }
            );
            if (bold) rPr.AppendChild(new Bold());

            return new TableCell(cellProps,
                new Paragraph(
                    new ParagraphProperties(
                        new Justification { Val = JustificationValues.Center }),
                    new Run(rPr,
                        new Text(text) { Space = SpaceProcessingModeValues.Preserve })));
        }

        private static T MakeWordBorder<T>() where T : BorderType, new() =>
            new T { Val = BorderValues.Single, Size = 4, Color = "CCCCCC", Space = 0 };

        private static Paragraph MakeWordPageBreak() =>
            new Paragraph(new Run(new Break { Type = BreakValues.Page }));

        private static void AddWordStyles(MainDocumentPart mainPart)
        {
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(
                new DocDefaults(
                    new RunPropertiesDefault(
                        new RunPropertiesBaseStyle(
                            new RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                            new WordFontSize { Val = "22" }
                        )
                    )
                )
            );
            stylesPart.Styles.Save();
        }
        private static void AskOpen(string filePath)
        {
            if (MessageBox.Show("Открыть файл?", "Готово",
                    MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                { FileName = filePath, UseShellExecute = true });
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
        public string Source { get; set; } = "xlsx";
    }

    public class JsonOrder
    {
        public int Id { get; set; }
        public string? OrderCode { get; set; }
        public string? CreatedDate { get; set; }
        public string? ClientCode { get; set; }
        public string? Services { get; set; }
        public string? Status { get; set; }
    }
}
