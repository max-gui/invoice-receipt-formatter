using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Ms = Microsoft.Office.Interop.Excel;

namespace automate_output
{
    class Program
    {
        static void Main(string[] args)
        {
            var MyApp = new Ms.Application();
            var MyBooks = default(Ms.Workbooks);
            var MyBook = default(Ms.Workbook);
            var MySheets = default(Ms.Sheets);
            var MySheet = default(Ms.Worksheet);
            var MyRange = default(Ms.Range);
            var lastRange = default(Ms.Range);
            try
            {
                MyApp.Visible = false;//System.Environment.CurrentDirectory
                MyApp.DisplayAlerts = false;
                MyBooks = MyApp.Workbooks;
                MyBook = MyBooks.Open($@"{System.Environment.CurrentDirectory}\output.xlsx");// DB_PATH);
                MySheets = MyBook.Sheets;
                MySheet = MySheets[1]; // Explict cast is not required here
                MyRange = MySheet.Cells;
                lastRange = MyRange.SpecialCells(Ms.XlCellType.xlCellTypeLastCell);
                var lastRow = lastRange.Row;
                lastRow = 2;


                string[] lines = System.IO.File.ReadAllLines($@"{System.Environment.CurrentDirectory}\input.txt", Encoding.GetEncoding("gb2312"));

                List<outputEntity> opeList = new List<outputEntity>();
                opeList.Clear();
                var nameTmp = string.Empty;
                foreach (var line in lines)
                {
                    if (line.StartsWith("~"))
                    {
                        nameTmp = line.TrimStart('~');
                        continue;
                    }
                    var moneyTemp = int.Parse(GetSubStringBetween(line, "'", "'"));
                    var date = DateTime.Parse(GetSubStringBetween(line, "*", "*"));
                    var onTime = DateTime.Parse(GetSubStringBetween(line, "@", "@"));
                    var emp = new outputEntity
                    {
                        NameIndex1 = nameTmp,
                        Data2 = date.ToShortDateString(),
                        OnTime3 = onTime.ToShortTimeString(),
                        DownTime4 = string.Empty,
                        Money5 = moneyTemp.ToString(),
                        WaitingTime6 = string.Empty,
                        Message7 = "上海"
                    };
                    opeList.Add(emp);

                    WriteToExcel(MySheet, ++lastRow, emp);
                }

                lastRow++;

                ParkingFee(MySheet, lastRow);

                MyBook.SaveAs($@"{System.Environment.CurrentDirectory}\{DateTime.Now.ToString("yyyyMM")}-output.xlsx");//, AccessMode: Ms.XlSaveAsAccessMode.xlNoChange);//, ConflictResolution: Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges);

                MyBook.Saved = true;
            }
            catch (Exception e)
            {
                MyBook.Saved = true;
                Console.WriteLine(e.Message);
                Console.ReadLine();
            }
            finally
            {
                MyApp.Quit();

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();

                releaseObject(lastRange);
                releaseObject(MyRange);
                releaseObject(MySheet);
                releaseObject(MySheets);
                releaseObject(MyBook);
                releaseObject(MyBooks);
                releaseObject(MyApp);

                //MyBooks = default(Ms.Workbooks);
                //var MyBook = default(Ms.Workbook);
                //var MySheets = default(Ms.Sheets);
                //var MySheet = default(Ms.Worksheet);
                //var MyRange = default(Ms.Range);
                //var lastRange = default(Ms.Range);
            }

        }

        private static void ParkingFee(Ms.Worksheet MySheet, int lastRow)
        {
            var mergedSheet = MySheet.get_Range($"A{lastRow}", $"G{lastRow}");
            MySheet.Cells[lastRow, 1] = "物业";
            MySheet.Cells[lastRow, 2] = "日期";
            MySheet.Cells[lastRow, 3] = "金额";
            MySheet.Cells[lastRow, 4] = "抬头";
            MySheet.get_Range($"D{lastRow}", $"F{lastRow}").Merge(false);
            var colorTmp = MySheet.Cells[3, 1].Interior.Color;
            mergedSheet.Interior.Color = colorTmp;
            lastRow++;

            mergedSheet = MySheet.get_Range($"D{lastRow}", $"F{lastRow + 1}");
            mergedSheet.Merge(false);
            mergedSheet.Value = "停车费抬头:上海希明电气技术有限公司";

            //mergedSheet.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
            //mergedSheet.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            mergedSheet.Font.Size = 9;

            MySheet.Cells[lastRow, 1] = "李莹怡";
            var dateTmp = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 20 + new Random().Next(0,5));
            MySheet.Cells[lastRow, 2] = dateTmp;
            MySheet.Cells[lastRow, 3] = 350;
            MySheet.Cells[lastRow, 7] = "上海";
        }

        public static void WriteToExcel(Ms.Worksheet MySheet, int lastRow, outputEntity emp)
        {
            try
            {
                MySheet.Cells[lastRow, 1] = emp.NameIndex1;
                MySheet.Cells[lastRow, 2] = emp.Data2;
                MySheet.Cells[lastRow, 3] = emp.OnTime3;
                MySheet.Cells[lastRow, 4] = emp.DownTime4;

                MySheet.Cells[lastRow, 5] = emp.Money5;

                MySheet.Cells[lastRow, 6] = emp.WaitingTime6;

                MySheet.Cells[lastRow, 7] = emp.Message7;
            }
            catch (Exception ex)
            { }

        }

        private static string GetSubStringBetween(string originString, string first, string last)
        {
            var start = originString.IndexOf(first);
            var end = originString.LastIndexOf(last);
            var startTmp = start + 1;
            var strRet = originString.Substring(startTmp, end - (startTmp));

            Console.WriteLine(strRet);
            return strRet;
        }

        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
