using System.Windows;
using Microsoft.Win32;
using System.Data;
using System.Windows.Input;

namespace ReadExcel_And_BindToDataGrid
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {
            //Open file dialog to get the location to the file
            OpenFileDialog openfile = new OpenFileDialog();
            openfile.DefaultExt = ".xlsx";
            openfile.Filter = "(.xlsx)|*.xlsx";

            var browsefile = openfile.ShowDialog();

            if (browsefile == true)
            {
                //visualize the path to the user
                txtFilePath.Text = openfile.FileName;

                //Get needed excel files
                var excelApp = new Microsoft.Office.Interop.Excel.Application();
                var excelBook = excelApp.Workbooks.Open(txtFilePath.Text.ToString(), 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                var excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(1); ;
                var excelRange = excelSheet.UsedRange;

                //Create DataTable where we will put the information from the excel file
                var dt = new DataTable();
                int excelColCount = excelRange.Columns.Count;

                //Set the column names
                for (int currCol = 1; currCol <= excelColCount; currCol++)
                {
                    string strColumn = (string)(excelRange.Cells[2, currCol] as Microsoft.Office.Interop.Excel.Range).Value2;
                    dt.Columns.Add(strColumn, typeof(string));
                }

                //Set the column info
                for (int currRow = 3; currRow <= excelRange.Rows.Count; currRow++)
                {
                    var strData = new string[excelRange.Columns.Count];
                    for (int currCol = 1; currCol <= excelRange.Columns.Count; currCol++)
                    {
                        strData[currCol - 1] = (string)(excelRange.Cells[currRow, currCol] as Microsoft.Office.Interop.Excel.Range).Value2;
                    }
                    dt.Rows.Add(strData);
                }

                //Sort the DataTable by Name ascending
                dt.DefaultView.Sort = "Name asc";

                //Display the information :)
                dtGrid.ItemsSource = dt.DefaultView;

                //Close excel conn
                excelBook.Close(true, null, null);
                excelApp.Quit();
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            //Allow for window drag
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                this.DragMove();
            }
        }
    }
}
