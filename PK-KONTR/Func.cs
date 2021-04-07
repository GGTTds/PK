using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows.Forms;

namespace PK_KONTR
{
    class Func
    {
        public static void Viz(string[] s)
        {
           
            var App = new Excel.Application();
            Excel.Workbook xlWB;
            try
            { 
            string L;
            StreamReader rr = new StreamReader("Put.txt");
            L = rr.ReadLine();
            //"E:/.../CreatEXcelQWWordOT/CreatEXcelQWWordOT/bin/Debug/Form.xlsx"
            L = L.Replace(@"\", "/");
            //string xlFileName = L;
            xlWB = App.Workbooks.Open(L);
            int StartIndex = 2;
            Excel.Worksheet worksheet1 = App.Worksheets["Данные"];
            Excel.Worksheet worksheet2 = App.Worksheets["Отчет"];
            int Ind = 0;
            int k = 0;
            //worksheet1.Cells[6][7] = "sfsfssf";
            int lastRow = worksheet1.Cells[1].SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                //MessageBox.Show($"Последняя строчка: {lastRow}");
                Excel.Range Head291 = worksheet2.Range[worksheet2.Cells[1][1], worksheet2.Cells[6][1]];
                Head291.Merge();
                worksheet2.Cells[1][1].Formula = "                                              Входной контроль Приход №Протокол проверки закладных из  от  № - 20 от 20г.";
                for (int i = 1; k <= Start.KolPov; i++)
            {
                    //Start.str[Ind].Equals(worksheet1.Cells[6][i].Formula)
                    //worksheet1.Cells[6][i].Formula == Start.str[Ind]
                    
                    int pp = StartIndex;
                    worksheet2.Columns.AutoFit();
                    //Rng = worksheet2.Cells[6].Find(Start.str[Ind], Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart); //осуществляем поиск на листе
                    if (Start.str[Ind].Contains(worksheet1.Cells[6][i].Formula)) 
                {
                    int st = i;
                    int sd = StartIndex;
                        //MessageBox.Show(Start.str[Ind].ToString()); 
                        worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][st];
                        worksheet2.Cells[6][StartIndex] = "Комментарий";
                            worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][st];
                            worksheet2.Cells[1][StartIndex + 1] = worksheet1.Cells[3][st];
                            Excel.Range Head21 = worksheet2.Range[worksheet2.Cells[2][StartIndex], worksheet2.Cells[5][StartIndex]];
                            Head21.Merge();
                        Excel.Range Head23 = worksheet2.Range[worksheet2.Cells[1][StartIndex], worksheet2.Cells[6][StartIndex]];
                        Head23.Interior.Color = Excel.XlRgbColor.rgbGray;
                        StartIndex += 1;
                        int po = StartIndex;
                        worksheet2.Cells[1][StartIndex] = "Поступило, шт";
                            worksheet2.Cells[2][StartIndex] = Start.str1[Ind];
                            Excel.Range Head211 = worksheet2.Range[worksheet2.Cells[2][StartIndex], worksheet2.Cells[5][StartIndex]];
                            Head211.Merge();
                            StartIndex += 1;
                            worksheet2.Cells[1][StartIndex] = worksheet1.Cells[5][st];
                            worksheet2.Cells[2][StartIndex] = worksheet1.Cells[6][st];
                            Excel.Range Head214 = worksheet2.Range[worksheet2.Cells[2][StartIndex], worksheet2.Cells[5][StartIndex]];
                            Head214.Merge();
                            StartIndex += 1;
                            worksheet2.Cells[1][StartIndex] = " ВидПроверки ";
                            worksheet2.Cells[2][StartIndex] = " Норма ";
                            worksheet2.Cells[3][StartIndex] = " Факт ";
                            worksheet2.Cells[4][StartIndex] = " Проверено, шт ";
                            worksheet2.Cells[5][StartIndex] = " Несоотв., шт ";
                            StartIndex += 1;
                            worksheet2.Cells[1][StartIndex] = worksheet1.Cells[7][st];
                            worksheet2.Cells[2][StartIndex] = worksheet1.Cells[8][st];
                            //Excel.Range Head27 = worksheet2.Range[worksheet2.Cells[2][StartIndex], worksheet2.Cells[6][StartIndex]];
                            //Head27.Merge();
                            StartIndex += 1;
                            worksheet2.Cells[1][StartIndex] = worksheet1.Cells[9][st];
                            worksheet2.Cells[2][StartIndex] = worksheet1.Cells[10][st];
                            for (int r = 11; r <= 50; r += 1)
                            {
                                if (worksheet1.Cells[r][st].value == null)
                                {
                                }
                                else
                                {
                                    StartIndex += 1;
                                    worksheet2.Cells[1][StartIndex] = worksheet1.Cells[r][st];
                                    worksheet2.Cells[2][StartIndex] = worksheet1.Cells[r += 1][st];
                                }
                            }
                        Excel.Range Head212 = worksheet2.Range[worksheet2.Cells[6][po], worksheet2.Cells[6][StartIndex]];
                        Head212.Merge();
                        Excel.Range RR1 = worksheet2.Range[worksheet2.Cells[1][pp], worksheet2.Cells[6][StartIndex]];
                        RR1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                        StartIndex += 1;
                       }

                    if (i == lastRow)
                    {
                        if (Ind <= Start.KolPov)
                        { Ind += 1; }

                        Start.KolPov -= 1;
                        i = 1;
                        k += 1;
                    }
                    
            }
            }
            catch
            {
                App.Quit();
                MessageBox.Show(" Ошибка!!! Перезапустите приложение");
            }

            App.Visible = true;        }
    }
}
