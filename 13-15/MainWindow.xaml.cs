using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms.DataVisualization.Charting;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;


namespace _13_15
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            ChartPayments.ChartAreas.Add(new ChartArea("Main"));
            var currentSeries = new Series("Платежи")
            {
                IsValueShownAsLabel = true
            };
            ChartPayments.Series.Add(currentSeries);

            CmbUser.ItemsSource = _context.User.ToList();
            CmbDiagram.ItemsSource = Enum.GetValues(typeof(SeriesChartType));
        }
        private Entities _context = new Entities();



        private void UpdateChart(object sender, SelectionChangedEventArgs e)
        {
            if (CmbUser.SelectedItem is User currentUser && CmbDiagram.SelectedItem is SeriesChartType currentType)
            {
                Series currentSeries = ChartPayments.Series.FirstOrDefault();
                currentSeries.ChartType = currentType;
                currentSeries.Points.Clear();

                var categoriesList = _context.Category.ToList();
                foreach (var category in categoriesList)
                {
                    currentSeries.Points.AddXY(category.Name,
                        _context.Payment.ToList().Where(u => u.User == currentUser
                        && u.Category == category).Sum(u => u.Price * u.Num));
                }
            }
        }


        private void WindowClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var result = MessageBox.Show("Вы точно хотите выйти из приложения?", "Выход", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                e.Cancel = false;
            }
            else
            {
                e.Cancel = true;
            }
        }

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            var allUsers = _context.User.ToList().OrderBy(u => u.FIO).ToList();
            var application = new Excel.Application();
            application.SheetsInNewWorkbook = allUsers.Count();
            Excel.Workbook workbook = application.Workbooks.Add(Type.Missing);
            for (int i = 0; i < allUsers.Count(); i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = application.Worksheets.Item[i + 1];
                worksheet.Name = allUsers[i].FIO;
                worksheet.Cells[1][startRowIndex] = "Дата платежа";
                worksheet.Cells[2][startRowIndex] = "Название";
                worksheet.Cells[3][startRowIndex] = "Стоимость";
                worksheet.Cells[4][startRowIndex] = "Количество";
                worksheet.Cells[5][startRowIndex] = "Сумма";
                Excel.Range columlHeaderRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[5][1]];
                columlHeaderRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                columlHeaderRange.Font.Bold = true;
                startRowIndex++;
                var userCategories = allUsers[i].Payment.OrderBy(u => u.Date).GroupBy(u => u.Category).OrderBy(u => u.Key.Name);
                double totalSum = 0; // Хранение общей суммы для пользователя
                foreach (var groupCategory in userCategories)
                {
                    Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][startRowIndex], worksheet.Cells[5][startRowIndex]];
                    headerRange.Merge();
                    headerRange.Value = groupCategory.Key.Name;
                    headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    headerRange.Font.Italic = true;
                    startRowIndex++;

                    foreach (var payment in groupCategory)
                    {
                        worksheet.Cells[1][startRowIndex] = payment.Date.ToString("dd.MM.yyyy");
                        worksheet.Cells[2][startRowIndex] = payment.Name;
                        worksheet.Cells[3][startRowIndex] = payment.Price;
                        (worksheet.Cells[3][startRowIndex] as Excel.Range).NumberFormat = "0.00";
                        worksheet.Cells[4][startRowIndex] = payment.Num;
                        worksheet.Cells[5][startRowIndex].Formula = $"=C{startRowIndex}*D{startRowIndex}";
                        (worksheet.Cells[5][startRowIndex] as Excel.Range).NumberFormat = "0.00";
                        startRowIndex++;
                    } // завершение цикла по платежам
                    Excel.Range sumRange = worksheet.Range[worksheet.Cells[1][startRowIndex], worksheet.Cells[4][startRowIndex]];
                    sumRange.Merge();
                    sumRange.Value = "ИТОГО:";
                    sumRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    worksheet.Cells[5][startRowIndex].Formula = $"=SUM(E{startRowIndex - groupCategory.Count()}:" + $"E{startRowIndex - 1})";
                    totalSum += worksheet.Cells[5][startRowIndex].Value; // Добавление суммы к общему итогу
                    sumRange.Font.Bold = worksheet.Cells[5][startRowIndex].Font.Bold = true;
                    startRowIndex++;
                    Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[5][startRowIndex - 1]];
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                }
                // Добавление строки "Общий итог"
                Excel.Range totalSumRange = worksheet.Range[worksheet.Cells[1][startRowIndex], worksheet.Cells[4][startRowIndex]];
                totalSumRange.Merge();
                totalSumRange.Value = "ОБЩИЙ ИТОГ:";
                totalSumRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                totalSumRange.Font.Bold = true;
                totalSumRange.Interior.Color = Excel.XlRgbColor.rgbRed; // Красный цвет заливки
                worksheet.Cells[5][startRowIndex].Value = totalSum;
                worksheet.Cells[5][startRowIndex].Font.Bold = true;
                worksheet.Cells[5][startRowIndex].Interior.Color = Excel.XlRgbColor.rgbRed; // Красный цвет заливки
                startRowIndex++;

                worksheet.Columns.AutoFit();
            }
            application.Visible = true;
        }

        private void ExportToWord_Click(object sender, RoutedEventArgs e)
        {
            var allUsers = _context.User.ToList();
            var allCategories = _context.Category.ToList();
            var application = new Word.Application();
            Word.Document document = application.Documents.Add();
            foreach (Word.Section section in document.Sections)
            {
                Word.HeaderFooter header = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                Word.Range headerRange = header.Range;
                headerRange.Text = DateTime.Now.ToString("dd.MM.yyyy");
                headerRange.Font.Name = "Times New Roman";
                headerRange.Font.Size = 12;
                headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }

            // Добавление номера страницы в нижний колонтитул
            foreach (Word.Section section in document.Sections)
            {
                Word.HeaderFooter footer = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                Word.Range footerRange = footer.Range;
                footerRange.Font.Name = "Times New Roman";
                footerRange.Font.Size = 12;
                footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                footerRange.Fields.Add(footerRange, Word.WdFieldType.wdFieldPage);
            }
            foreach (var user in allUsers)
            {
                Word.Paragraph userParagraph = document.Paragraphs.Add();
                Word.Range userRange = userParagraph.Range;
                userRange.Text = user.FIO;
                userParagraph.set_Style("Заголовок");
                userRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                userRange.InsertParagraphAfter();
                document.Paragraphs.Add();
                Word.Paragraph tableParagraph = document.Paragraphs.Add();
                Word.Range tableRange = tableParagraph.Range;
                Word.Table paymentsTable = document.Tables.Add(tableRange, allCategories.Count() + 1, 2);
                paymentsTable.Borders.InsideLineStyle = paymentsTable.Borders.OutsideLineStyle =
                         Word.WdLineStyle.wdLineStyleSingle;
                paymentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                Word.Range cellRange;

                cellRange = paymentsTable.Cell(1, 1).Range;
                cellRange.Text = "Категория";
                cellRange = paymentsTable.Cell(1, 2).Range;
                cellRange.Text = "Сумма расходов";

                paymentsTable.Rows[1].Range.Font.Name = "Times New Roman";
                paymentsTable.Rows[1].Range.Font.Size = 14;
                paymentsTable.Rows[1].Range.Bold = 1;
                paymentsTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                for (int i = 0; i < allCategories.Count(); i++)
                {
                    var currentCategory = allCategories[i];
                    cellRange = paymentsTable.Cell(i + 2, 1).Range;
                    cellRange.Text = currentCategory.Name;
                    cellRange.Font.Name = "Times New Roman";
                    cellRange.Font.Size = 12;

                    cellRange = paymentsTable.Cell(i + 2, 2).Range;
                    cellRange.Text = user.Payment.ToList().
               Where(u => u.Category == currentCategory).Sum(u => u.Num * u.Price).ToString("N2") + " руб.";
                    cellRange.Font.Name = "Times New Roman";
                    cellRange.Font.Size = 12;
                } //завершение цикла по строкам таблицы
                document.Paragraphs.Add(); //пустая строка

                Payment maxPayment = user.Payment.OrderByDescending(u => u.Price * u.Num).FirstOrDefault();
                if (maxPayment != null)
                {
                    Word.Paragraph maxPaymentParagraph = document.Paragraphs.Add();
                    Word.Range maxPaymentRange = maxPaymentParagraph.Range;
                    maxPaymentRange.Text = $"Самый дорогостоящий платеж - {maxPayment.Name} за {(maxPayment.Price * maxPayment.Num).ToString("N2")} " + $"руб. от {maxPayment.Date.ToString("dd.MM.yyyy")}";
                    maxPaymentParagraph.set_Style("Подзаголовок");
                    maxPaymentRange.Font.Color = Word.WdColor.wdColorDarkRed;
                    maxPaymentRange.InsertParagraphAfter();
                }
                document.Paragraphs.Add(); //пустая строка

                Payment minPayment = user.Payment.OrderBy(u => u.Price * u.Num).FirstOrDefault();
                if (maxPayment != null)
                {
                    Word.Paragraph minPaymentParagraph = document.Paragraphs.Add();
                    Word.Range minPaymentRange = minPaymentParagraph.Range;
                    minPaymentRange.Text = $"Самый дешевый платеж - {minPayment.Name} за {(minPayment.Price * minPayment.Num).ToString("N2")} " + $"руб. от {minPayment.Date.ToString("dd.MM.yyyy")}";
                    minPaymentParagraph.set_Style("Подзаголовок");
                    minPaymentRange.Font.Color = Word.WdColor.wdColorDarkGreen;
                    minPaymentRange.InsertParagraphAfter();
                }

                if (user != allUsers.LastOrDefault()) document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);

                application.Visible = true;

                document.SaveAs2(@"D:\Payments.docx"); document.SaveAs2(@"D:\Payments.pdf", Word.WdExportFormat.wdExportFormatPDF);
            } //завершение цикла по пользователям

        }
    }        
}
