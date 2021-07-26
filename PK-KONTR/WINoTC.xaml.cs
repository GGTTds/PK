using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
using System.Threading.Tasks;
using System.Windows;
//using System.Windows.Controls;
//using System.Windows.Data;
//using System.Windows.Documents;
using System.Windows.Input;
//using System.Windows.Media;
//using System.Windows.Media.Imaging;
//using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace PK_KONTR
{
    public partial class WINoTC : Window
    {
        public WINoTC()
        {
            InitializeComponent();
            Start.KolPov = 0;
            Start.str = new string[1000];
            Start.str1 = new string[1000];
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        { GetFunc(); }

        private void WTO_KeyUp(object sender, KeyEventArgs e)
        {
            if(e.Key.Equals(Key.Enter))
            { GetFunc(); }
            if(e.Key.Equals(Key.Escape))
            { MainWindow ww = new MainWindow(); ww.Show(); this.Close(); }
        }
    
    
    public void GetFunc() 
        {
            
            int.TryParse(na.Text, out Start.Na4);
            int.TryParse(ko.Text, out Start.Kon4);
            if (Start.Na4 > Start.Kon4)
            {
                MessageBox.Show("Вы ввели некоректное значение", "Ошибка");
            }
            else
            {

                try
                {
                    string g;
                    var App = new Excel.Application();
                    Excel.Workbook xlWB;
                    string L;
                    StreamReader rr = new StreamReader("PutHH.txt");
                    L = rr.ReadLine();
                    //"E:/.../CreatEXcelQWWordOT/CreatEXcelQWWordOT/bin/Debug/Form.xlsx"
                    L = L.Replace(@"\", "/");
                    xlWB = App.Workbooks.Open(L);
                    int InFor = 0;
                    Excel.Worksheet worksheet2 = App.Worksheets["Лист1"];
                    for (int i = Start.Na4; i <= Start.Kon4; i++)
                    {
                        g = worksheet2.Cells[3][i].Formula;
                        g.Replace(@" ", "");
                        Start.str[InFor] = g;
                        Start.str1[InFor] = worksheet2.Cells[5][i].Formula;
                        InFor += 1;
                        Start.KolPov += 1;
                    }
                    Start.KolPov -= 1;
                    xlWB.Close(false,false,false);
                    App.Application.Quit();
                    App = null;
                    xlWB = null;
                    worksheet2 = null;
                    GC.Collect();
                    Func.Viz(Start.str);
                    WINoTC ww = new WINoTC();
                    ww.Show();
                    this.Close();
                }
                catch
                {
                    WINoTC ww = new WINoTC();
                    ww.Show();
                    this.Close();
                    MessageBox.Show(" Ошибка, введены не верные данные");
                }
            }
        }
        private void na_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !(Char.IsDigit(e.Text, 0));
        }
    }
}
