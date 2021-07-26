using Microsoft.Win32;
using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
using System.Threading.Tasks;
using System.Windows;
//using System.Windows.Controls;
//using System.Windows.Data;
//using System.Windows.Documents;
//using System.Windows.Input;
//using System.Windows.Media;
//using System.Windows.Media.Imaging;
//using System.Windows.Navigation;
//using System.Windows.Shapes;
using System.Windows.Forms;
using OpenFileDialog = System.Windows.Forms.OpenFileDialog;
using MessageBox = System.Windows.Forms.MessageBox;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace PK_KONTR
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            StreamReader rr = new StreamReader("Put.txt");
            Start.PutinBB = rr.ReadLine();
            rr.Close();
            Fail();
        }

        private void CreaOtch_Click(object sender, RoutedEventArgs e)
        {
            WINoTC WW = new WINoTC();
            WW.Show();
            this.Close();
        }

        private void Edit_Click(object sender, RoutedEventArgs e)
        {
            var App = new Excel.Application();
            Excel.Workbook xlWB;
            string L;
            StreamReader rr = new StreamReader("Put.txt");
            L = rr.ReadLine();
            L = L.Replace(@"\", "/");
            string xlFileName = L;
            xlWB = App.Workbooks.Open(L);
            App.Visible = true;
        }
        private void Add_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string str = dialog.FileName;
                Start.PutinBB = str;
                StreamWriter ss = new StreamWriter("Put.txt");
                ss.WriteLine(Start.PutinBB.ToString());
                ss.Close();
                DialogResult dialogResult = MessageBox.Show("Путь к файлу задан", "Файл", MessageBoxButtons.OK);
                if (dialogResult == System.Windows.Forms.DialogResult.OK) { Fail(); }
            }
        }
        public void Fail()
        {
           
            if ( Start.PutinBB != null)
            { fr.Content = " Выбран "; }
            else
            { fr.Content = " Не выбран"; }
        }
        private void Add_Copy_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string str = dialog.FileName;
                Start.PutinBB = str;
                StreamWriter ss = new StreamWriter("PutHH.txt");
                ss.WriteLine(Start.PutinBB.ToString());
                ss.Close();
                DialogResult dialogResult = MessageBox.Show("Путь к файлу задан", "Файл", MessageBoxButtons.OK);
                if (dialogResult == System.Windows.Forms.DialogResult.OK)
                { MessageBox.Show(" Файл выбран!"); }
            }
        }
    }
}
