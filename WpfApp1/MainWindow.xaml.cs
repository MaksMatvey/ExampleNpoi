using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;


namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            InitializeComponent();
            this.RowBox.PreviewTextInput += new TextCompositionEventHandler(RowBox_PreviewTextInput);//For numbers of Row
            this.ColumnBox.PreviewTextInput += new TextCompositionEventHandler(ColumnBox_PreviewTextInput);//For letters of Column

        }

        string textPathBox = @"C:\Users\Desktop\Новая папка";

        public void RowBox_PreviewTextInput(object sender, TextCompositionEventArgs e)//For numbers of Row
        {
            if (!Char.IsDigit(e.Text, 0))
            {
                e.Handled = true;
            }
        }

        public void ColumnBox_PreviewTextInput(object sender, TextCompositionEventArgs e)//For letters of Column
        {
            if (!Char.IsUpper(e.Text, 0))
            {
                e.Handled = true;
            }
        }

        public void ResetValueTextBox()
        {
            StringBox.Text = "";
            ColumnBox.Text = "";
            RowBox.Text = "";
            PathBox.Text = "";
            xlsRButton.IsChecked = false;
            xlsxRButton.IsChecked = false;
        }


        private void PathBox_GotFocus(object sender, RoutedEventArgs e)
        {

            if (PathBox.Text == textPathBox)
            {
                PathBox.Text = null;
            };
            PathBox.Foreground = System.Windows.Media.Brushes.Black;
        }

        private void PathBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (PathBox.Text == "")
            {
                PathBox.Foreground = System.Windows.Media.Brushes.Gray;
                PathBox.Text = textPathBox;
            }
        }

        public static XSSFWorkbook WorkBookXlsx = new XSSFWorkbook();
        XSSFSheet SheetXlsx = (XSSFSheet)WorkBookXlsx.CreateSheet("Sheet 1");
        public static HSSFWorkbook WorkBookXls = new HSSFWorkbook();
        HSSFSheet SheetXls = (HSSFSheet)WorkBookXls.CreateSheet("Sheet 1");
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            { 
            if (xlsxRButton.IsChecked == true)
            {
                    textPathBox = PathBox.Text + @"\result11.xlsx";
                    ExcelManager.CreateChangeExcelFile(SheetXlsx, WorkBookXlsx, RowBox.Text, ColumnBox.Text, StringBox.Text, textPathBox);
                    ResetValueTextBox();
                    MessageBox.Show("Data have written to file ");
            }
            if (xlsRButton.IsChecked == true)
            {
                    textPathBox = PathBox.Text + @"\result11.xls";
                    ExcelManager.CreateChangeExcelFile(SheetXls, WorkBookXls, RowBox.Text, ColumnBox.Text, StringBox.Text, textPathBox);
                    ResetValueTextBox();
                    MessageBox.Show("Data have written to file ");
            }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
    }
}
