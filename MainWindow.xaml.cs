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
        private int productAmounts = 51;
        //private string itemToChange;
        private int itemToChange;
        //will be set after button click item number = row number
        string productXLS = "A";
        string priceXLS = "B";
        string itemName;

        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = this;
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Product Sheet");

                for (int i = 1; i <= productAmounts; i++)
                {
                    worksheet.Cell(productXLS + i).Value = "EmptyItem" + i;
                    worksheet.Column(productXLS).AdjustToContents();
                    worksheet.Cell(priceXLS + i).Value = "0";
                    worksheet.Cell(priceXLS + i).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                }
                workbook.SaveAs("ProductTest.xlsx");
            }

            var workBook = new XLWorkbook("ProductTest.xlsx");
            var workSheet = workBook.Worksheet("Product Sheet");
            //workSheet.Cell(productXLS + itemToChange).Value = "New Item";
            //workBook.Save();
            for (int i = 1; i <= productAmounts; i++)
            {
                var data = workSheet.Cell(productXLS + i).GetValue<string>();
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


            //workSheet.Cell(productXLS + itemToChange).Value = "New Item";
            //workBook.Save();
        }

        private void btn1_Click(object sender, RoutedEventArgs e)
        {
            itemToChange = 1;
            SetProduct();
        }

        private void btn2_Click(object sender, RoutedEventArgs e)
        {
            itemToChange = 2;
            SetProduct();
        }
        private void SetProduct()
        {
            var workBook = new XLWorkbook("ProductTest.xlsx");
            var workSheet = workBook.Worksheet("Product Sheet");
            workSheet.Cell("A1").Value = "Pizza";
            workSheet.Cell("A2").Value = "Burger";
            var btnName = "btn";
            var data = workSheet.Cell("A" + itemToChange).GetValue<string>();
            
            for (int i = 1; i <= productAmounts; i++)
            {
                if (itemToChange == i)
                {
                    btnName = btnName + i;
                    foreach (UIElement item in grTest.Children)
                    {
                        if (item.GetType() == typeof(Button))
                        {
                            Button newButton = (Button)item;
                            if (newButton.Name == btnName)
                            {
                                newButton.Content = data;
                            }
                        }
                    }
                }
            }
        }
    }
}

