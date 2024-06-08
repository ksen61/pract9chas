using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
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
using System.Windows.Shapes;

namespace WpfApp19
{
    /// <summary>
    /// Логика взаимодействия для OpenExcel.xaml
    /// </summary>
    public partial class OpenExcel : Window
    {
        private string path;
        public OpenExcel(string path)
        {
            InitializeComponent();
            this.path = path;
            Workbook wb = new Workbook();
            wb.LoadFromFile(path);
            Worksheet sheet = wb.Worksheets[0];
            CellRange locatedRange = sheet.AllocatedRange;

            var dataTable = sheet.ExportDataTable(locatedRange, true);

            grid.ItemsSource = dataTable.DefaultView;
        }

        private void AddColumn_Click(object sender, RoutedEventArgs e)
        {
            Workbook wb = new Workbook();
            wb.LoadFromFile(path);
            Worksheet sheet = wb.Worksheets[0];
            CellRange locatedRange = sheet.AllocatedRange;

            var dataTable = sheet.ExportDataTable(locatedRange, true);
            DataColumn newColumn = new DataColumn();
            newColumn.ColumnName = NameColumn.Text;
            dataTable.Columns.Add(newColumn);
            grid.ItemsSource = dataTable.DefaultView;
        }

        private void SaveBtm_Click(object sender, RoutedEventArgs e)
        {
            Save();
        }

        private void SendBtm_Click(object sender, RoutedEventArgs e)
        {
            Save();
            Sending sending = new Sending(path);
            sending.ShowDialog();
        }

        private void Save()
        {
            var dataTable = grid.ItemsSource as DataView;

            Workbook wb = new Workbook();
            wb.Worksheets.Clear();
            Worksheet sheet = wb.Worksheets.Add("Лист 1");
            sheet.InsertDataView(dataTable, true, 1, 1);

            wb.SaveToFile(path, FileFormat.Version2016); ;
        }
    }
}
