using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;
using System.Data.Entity.Validation;
using System.IO;
using Path = System.IO.Path;
using System.Text.Json;
using static System.Net.Mime.MediaTypeNames;


namespace Template_4335.Windows
{
    /// <summary>
    /// Логика взаимодействия для Mukhametzyanova_M_4335.xaml
    /// </summary>
    public partial class Mukhametzyanova_M_4335 : System.Windows.Window
    {
        public Mukhametzyanova_M_4335()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //OpenFileDialog ofd = new OpenFileDialog()
            //{
            //    DefaultExt = "*.xls;*.xlsx",
            //    Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
            //    Title = "Выберите файл базы данных"
            //};
            //if (!(ofd.ShowDialog() == true))
            //    return;
            //string[,] list;
            //Excel.Application ObjWorkExcel = new Excel.Application();
            //Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            //Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            //var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            //int _columns = (int)lastCell.Column;
            //int _rows = (int)lastCell.Row;
            //list = new string[_rows, _columns];
            //for (int j = 0; j < _columns; j++)
            //    for (int i = 0; i < _rows; i++)
            //        list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            //ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            //ObjWorkExcel.Quit();
            //GC.Collect();

            //using (UslugaEntities usersEntities = new UslugaEntities())
            //{
            //    for (int i = 0; i < _rows; i++)
            //    {
            //        usersEntities.Uslugas.Add(new Uslugas()
            //        {
            //            Name = list[i, 1],
            //            Type = list[i, 2],
            //            Cost = list[i, 4]
            //        });
            //    }
            //    usersEntities.SaveChanges();
            //}

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            //List<Uslugas3> allStudents;

            //using (UslugaEntities1 usersEntities = new UslugaEntities1())
            //{
            //    allStudents = usersEntities.Uslugas3.ToList().OrderBy(s => s.Id).ToList();
            //}
            //var app = new Excel.Application();
            //app.SheetsInNewWorkbook = allStudents.Count();
            //Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            //for (int i = 1; i < 4; i++)
            //{
            //    int startRowIndex = 1;
            //    Excel.Worksheet worksheet = app.Worksheets.Item[i];
            //    worksheet.Name = "Категория " + Convert.ToString(i);
            //    worksheet.Cells[1][startRowIndex] = "Порядковый номер";
            //    worksheet.Cells[2][startRowIndex] = "Название";
            //    worksheet.Cells[3][startRowIndex] = "Тип";
            //    worksheet.Cells[4][startRowIndex] = "Стоимость";
            //    startRowIndex++;

            //    foreach (var usluga in allStudents)
            //    {
                    
            //            string tip = "";

            //            if (Convert.ToInt32(usluga.Cost) <= 250)

            //            {
            //                tip = "Категория 1"; }

            //            if (Convert.ToInt32(usluga.Cost) <= 800 && Convert.ToInt32(usluga.Cost) > 250)
            //            { tip = "Категория 2"; }

            //            if (Convert.ToInt32(usluga.Cost) > 800) { tip = "Категория 3"; }

            //            if (tip == worksheet.Name)
            //            {
            //                worksheet.Cells[1][startRowIndex] = usluga.Id;
            //                worksheet.Cells[2][startRowIndex] = usluga.Name;
            //                worksheet.Cells[3][startRowIndex] = usluga.Type;
            //                worksheet.Cells[4][startRowIndex] = usluga.Cost;
            //                startRowIndex++;
            //            }
            //        }

                

            //    worksheet.Columns.AutoFit();
            //}
            //app.Visible = true;

        }

        private async void ExportJson_Click(object sender, RoutedEventArgs e)
        {
               var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "Mukhametzyanova_M_4335", "data", "1.json");

                using (var fileStream = new FileStream(path, FileMode.Open))
                using (var db = new UslugaEntities2())
                {
                    var Uslugas3 = await JsonSerializer.DeserializeAsync<List<Uslugas3>>(fileStream);

                    foreach (Uslugas3 item in Uslugas3)
                    {
                        var service = new Uslugas3
                        {
                            IdServices = item.IdServices,
                            NameServices = item.NameServices,
                            TypeOfService = item.TypeOfService,
                            CodeService  = item.CodeService,
                            Cost = item.Cost
                        };

                        db.Uslugas3.Add(service);
                    }
                    try
                    {
                    db.SaveChanges();
                    //datagrid.Items.Refresh();
                    MessageBox.Show("Данные импортированы успешно!",
                                        "Внимание!",
                                        MessageBoxButton.OK,
                                        MessageBoxImage.Information);
                        db.SaveChanges();
                    }
                    catch (DbEntityValidationException ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }

        

        private void ImportJson_Click(object sender, RoutedEventArgs e)
        {
            #region Категории по возрасту

            var firstRangePrice = new List<Uslugas3>();
            var secondRangePrice = new List<Uslugas3>();
            var thirdRangePrice = new List<Uslugas3>();

            #endregion

            using (var db = new UslugaEntities2())
            {
                #region Сортировка по цене

                firstRangePrice = db.Uslugas3.ToList()
                                             .Where(fR => Convert.ToInt32(fR.Cost) >= 0 &&
                                                          Convert.ToInt32(fR.Cost) <= 350)
                                             .GroupBy(fR => fR.NameServices)
                                             .SelectMany(fR => fR)
                                             .ToList();

                secondRangePrice = db.Uslugas3.ToList()
                                              .Where(sR => Convert.ToInt32(sR.Cost) >= 250 &&
                                                         Convert.ToInt32(sR.Cost) <= 800)
                                              .GroupBy(sR => sR.NameServices)
                                              .SelectMany(sR => sR)
                                              .ToList();

                thirdRangePrice = db.Uslugas3.ToList()
                                             .Where(tR => Convert.ToInt32(tR.Cost) >= 800)
                                             .GroupBy(tR => tR.NameServices)
                                             .SelectMany(tR => tR)
                                             .ToList();

                #endregion

                #region Создание Word 

                var app = new Microsoft.Office.Interop.Word.Application();
                var document = app.Documents.Add();

                #endregion

                #region Создание параграфов

                #region Создание таблицы для первой категории

                for (var i = 0; i < 1; i++)
                {
                    var paragraph = document.Paragraphs.Add();
                    var range = paragraph.Range;
                    range.Text = "Категория цен от 0 до 350 рублей";
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    var tableParagraph = document.Paragraphs.Add();
                    var tableRange = tableParagraph.Range;
                    var priceCategories = document.Tables.Add(tableRange, firstRangePrice.Count() + 1, 4);

                    priceCategories.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                    priceCategories.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleDot;
                    priceCategories.Range.Cells.VerticalAlignment = (Microsoft.Office.Interop.Word.WdCellVerticalAlignment)Microsoft.Office.Interop.Word.WdVerticalAlignment.wdAlignVerticalCenter;

                    Microsoft.Office.Interop.Word.Range cellRange;
                    cellRange = priceCategories.Cell(1, 1).Range;
                    cellRange.Text = "Идентификатор";
                    cellRange = priceCategories.Cell(1, 2).Range;
                    cellRange.Text = "Наименование услуги";
                    cellRange = priceCategories.Cell(1, 3).Range;
                    cellRange.Text = "Тип услуги";
                    cellRange = priceCategories.Cell(1, 4).Range;
                    cellRange.Text = "Цена услуги";
                    priceCategories.Rows[1].Range.Bold = 1;
                    priceCategories.Rows[1].Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    var count = 1;
                    foreach (var item in firstRangePrice)
                    {
                        cellRange = priceCategories.Cell(count + 1, 1).Range;
                        cellRange.Text = Convert.ToString(item.IdServices);
                        cellRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = priceCategories.Cell(count + 1, 2).Range;
                        cellRange.Text = item.NameServices;
                        cellRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = priceCategories.Cell(count + 1, 3).Range;
                        cellRange.Text = item.TypeOfService;
                        cellRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = priceCategories.Cell(count + 1, 4).Range;
                        cellRange.Text = Convert.ToString(item.Cost);
                        cellRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        count++;
                    }
                }

                #endregion

                #region Создание таблицы для второй категории

                for (var i = 0; i < 1; i++)
                {
                    var paragraph = document.Paragraphs.Add();
                    var range = paragraph.Range;
                    range.Text = "Категория цены от 250 до 800 рублей";
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    var tableParagraph = document.Paragraphs.Add();
                    var tableRange = tableParagraph.Range;
                    var priceCategories = document.Tables.Add(tableRange, secondRangePrice.Count() + 1, 4);

                    priceCategories.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                    priceCategories.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleDot;
                    priceCategories.Range.Cells.VerticalAlignment = (Microsoft.Office.Interop.Word.WdCellVerticalAlignment)Microsoft.Office.Interop.Word.WdVerticalAlignment.wdAlignVerticalCenter;

                    Microsoft.Office.Interop.Word.Range cellRange;
                    cellRange = priceCategories.Cell(1, 1).Range;
                    cellRange.Text = "Идентификатор";
                    cellRange = priceCategories.Cell(1, 2).Range;
                    cellRange.Text = "Наименование услуги";
                    cellRange = priceCategories.Cell(1, 3).Range;
                    cellRange.Text = "Тип услуги";
                    cellRange = priceCategories.Cell(1, 4).Range;
                    cellRange.Text = "Цена услуги";
                    priceCategories.Rows[1].Range.Bold = 1;
                    priceCategories.Rows[1].Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    var count = 1;
                    foreach (var item in secondRangePrice)
                    {
                        cellRange = priceCategories.Cell(count + 1, 1).Range;
                        cellRange.Text = Convert.ToString(item.IdServices);
                        cellRange.ParagraphFormat.Alignment =   Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = priceCategories.Cell(count + 1, 2).Range;
                        cellRange.Text = item.NameServices;
                        cellRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = priceCategories.Cell(count + 1, 3).Range;
                        cellRange.Text = item.TypeOfService;
                        cellRange.ParagraphFormat.Alignment =   Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = priceCategories.Cell(count + 1, 4).Range;
                        cellRange.Text = Convert.ToString(item.Cost);
                        cellRange.ParagraphFormat.Alignment =   Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        count++;
                    }
                }

                #endregion

                #region Создание таблицы для третьей категории

                for (int i = 0; i < 1; i++)
                {
                    var paragraph = document.Paragraphs.Add();
                    var range = paragraph.Range;
                    range.Text = "Категория цен от 800 рублей";
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    var tableParagraph = document.Paragraphs.Add();
                    var tableRange = tableParagraph.Range;
                    var priceCategories = document.Tables.Add(tableRange, thirdRangePrice.Count() + 1, 4);

                    priceCategories.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                    priceCategories.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleDot;
                    priceCategories.Range.Cells.VerticalAlignment = (Microsoft.Office.Interop.Word.WdCellVerticalAlignment)Microsoft.Office.Interop.Word.WdVerticalAlignment.wdAlignVerticalCenter;

                    Microsoft.Office.Interop.Word.Range cellRange;
                    cellRange = priceCategories.Cell(1, 1).Range;
                    cellRange.Text = "Идентификатор";
                    cellRange = priceCategories.Cell(1, 2).Range;
                    cellRange.Text = "Наименование услуги";
                    cellRange = priceCategories.Cell(1, 3).Range;
                    cellRange.Text = "Тип услуги";
                    cellRange = priceCategories.Cell(1, 4).Range;
                    cellRange.Text = "Цена услуги";
                    priceCategories.Rows[1].Range.Bold = 1;
                    priceCategories.Rows[1].Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    var count = 1;
                    foreach (var item in thirdRangePrice)
                    {
                        cellRange = priceCategories.Cell(count + 1, 1).Range;
                        cellRange.Text = Convert.ToString(item.IdServices);
                        cellRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = priceCategories.Cell(count + 1, 2).Range;
                        cellRange.Text = item.NameServices;
                        cellRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = priceCategories.Cell(count + 1, 3).Range;
                        cellRange.Text = item.TypeOfService;
                        cellRange.ParagraphFormat.Alignment =   Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = priceCategories.Cell(count + 1, 4).Range;
                        cellRange.Text = Convert.ToString(item.Cost);
                        cellRange.ParagraphFormat.Alignment =   Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        count++;
                    }
                }

                #endregion

                app.Visible = true;

                #endregion
            }
        }

       
    }
}


