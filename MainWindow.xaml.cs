using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ClosedXML;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.VariantTypes;

namespace XlsSaveFile
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            InitializeComponent();
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Product Sheet");

                for (int i = 1; i <= 10; i++)
                {
                    string productXLS = "A";
                    string priceXLS = "B";
                    worksheet.Cell(productXLS + i).Value = "EmptyItem" + i;
                    worksheet.Column(productXLS).AdjustToContents();
                    worksheet.Cell(priceXLS + i).Value = "0";
                    worksheet.Cell(priceXLS + i).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                }
                workbook.SaveAs("ProductTest.xlsx");
            }
            var workBook = new XLWorkbook("ProductTest.xlsx");
            var workSheet = workBook.Worksheet("Product Sheet");
            //var ws1 = wbook.Worksheet(1);
            for (int i = 1; i <= 10; i++)
            {
                string productXLS = "A";
                var data = workSheet.Cell(productXLS + i).GetValue<string>();
                //Debug.WriteLine(data);


                string setLabel = "lbl" + i;
                foreach (UIElement item in grTest.Children)
                {
                    if (item.GetType() == typeof(TextBlock))
                    {
                        TextBlock txtBlock = (TextBlock)item;
                        if (txtBlock.Name == setLabel)
                        {
                            txtBlock.Text = data;
                        }
                    }
                }
            }
        }

    }
}

