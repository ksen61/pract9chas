using Microsoft.WindowsAPICodePack.Dialogs;
using Spire.Additions.Xps.Schema;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
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
    /// Логика взаимодействия для CreateExcel.xaml
    /// </summary>
    public partial class CreateExcel : Window
    {
        private string path;
        string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        DataTable dataTable1 = new DataTable();

        public CreateExcel()
        {
            InitializeComponent();
            grid.ItemsSource = dataTable1.DefaultView;
        }

        private void SaveBtm_Click(object sender, RoutedEventArgs e)
        {
            Save();
        }

        private void AddColumn_Click(object sender, RoutedEventArgs e)
        {
            DataColumn newColumn = new DataColumn();
            newColumn.ColumnName = NameColumn.Text;
            dataTable1.Columns.Add(newColumn);
            grid.ItemsSource = null;
            grid.ItemsSource = dataTable1.DefaultView;
        }

        private void SendBtm_Click(object sender, RoutedEventArgs e)
        {
            Save();
            Sending sending = new Sending(path);
            sending.ShowDialog();
        }
        private void Save()
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            CommonFileDialogResult result = dialog.ShowDialog();
            if (result == CommonFileDialogResult.Ok)
            {
                string selectedPath = dialog.FileName;

                var dataTable = grid.ItemsSource as DataView;
                Workbook wb = new Workbook();
                wb.Worksheets.Clear();
                Worksheet sheet = wb.Worksheets.Add("Лист 1");
                sheet.InsertDataView(dataTable, true, 1, 1);

                string filePath = selectedPath + "/" + NameFile.Text + ".xlsx";

                wb.SaveToFile(filePath, FileFormat.Version2016); ;

                MessageBox.Show("Данные сохранены в файл Excel. По этому пути: " + filePath);
            }
        }
    }
}