using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfApp19
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void CreateExcel_Click(object sender, RoutedEventArgs e)
        {
            CreateExcel exelCreateOkno = new CreateExcel();
            exelCreateOkno.Show();
            Close();
        }

        private void OpenExcel_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            CommonFileDialogResult result = dialog.ShowDialog();
            if (result == CommonFileDialogResult.Ok)
            {
                if (dialog.FileName.Contains(".xlsx"))
                {
                    string path = dialog.FileName;
                    OpenExcel openExcel = new OpenExcel(path);
                    openExcel.Show();
                    Close();
                }
                else
                {
                    MessageBox.Show("Это не эксель");
                }
            }
        }
    }
}
