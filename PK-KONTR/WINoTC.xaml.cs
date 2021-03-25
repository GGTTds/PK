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
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace PK_KONTR
{
    /// <summary>
    /// Логика взаимодействия для WINoTC.xaml
    /// </summary>
    public partial class WINoTC : Window
    {
        public WINoTC()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
           
                Start.Na4 = Convert.ToInt32(na.Text);
                Start.Kon4 = Convert.ToInt32(ko.Text);
                var App = new Excel.Application();
                Excel.Workbook xlWB;
            try
            { 
                string L;
                StreamReader rr = new StreamReader("PutHH.txt");
                L = rr.ReadLine();
                //"E:/.../CreatEXcelQWWordOT/CreatEXcelQWWordOT/bin/Debug/Form.xlsx"
                L = L.Replace(@"\", "/");
                xlWB = App.Workbooks.Open(L);
                int InFor = 0;
                Excel.Worksheet worksheet2 = App.Worksheets["Лист1"];
                //MessageBox.Show(Start.Na4.ToString());
                for (int i = Start.Na4; i <= Start.Kon4; i++)
                {

                    Start.str[InFor] = worksheet2.Cells[3][i].Formula;
                    InFor += 1;
                    Start.KolPov += 1;
                }

                App.Quit();
                Func.Viz(Start.str);
            }
            catch 
            {
                App.Quit();
                MessageBox.Show(" Ошибка, перезапустите приложение");
            }
            //MessageBox.Show(Start.str.LongLength.ToString());
            //    MessageBox.Show(Start.str[0].ToString());
            //    MessageBox.Show(Start.str[1].ToString());
            //    MessageBox.Show(Start.str[2].ToString());
            //    MessageBox.Show(Start.str[3].ToString());
        }

    }
}
