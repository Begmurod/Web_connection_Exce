
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;

namespace WebApplication2.Models
{
    public class ClassTest
    {
        public double B5 { get; set; }
        public double B6 { get; set; }
        public double B7 { get; set; }
        public double B9 { get; set; }
        public double B10 { get; set; }
        public double B11 { get; set; }
        public double B13 { get; set; }
        public double B14{ get; set; }
        public double B15{ get; set; }
        public double B17 { get; set; }
        public double B18 { get; set; }
        public double B19 { get; set; }
        public double B20 { get; set; }
        public double B21 { get; set; }
        public double B22 { get; set; }
        public double B23 { get; set; }
        public double B24 { get; set; }

        public void ExcelCon()
        {
            Excel.Application objExcel = null;
            Excel.Workbook WorkBook = null;

            try
            {
                objExcel = new Excel.Application();

                ////если надо показать Excel-файл
                objExcel.ScreenUpdating = true;
                objExcel.WindowState = Excel.XlWindowState.xlMaximized;
                objExcel.Visible = true;
                objExcel.DisplayAlerts = true;

                string fileName = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "OptTolStenki.xlsm");

                WorkBook = objExcel.Workbooks.Open(fileName, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                Excel.Worksheet WorkSheet = (Excel.Worksheet)WorkBook.Sheets["Calc"];
                WorkSheet.Range["B5"].Value2 = Convert.ToString(B5);
                WorkSheet.Range["B6"].Value2 = Convert.ToString(B6);
                WorkSheet.Range["B7"].Value2 = Convert.ToString(B7);
                WorkSheet.Range["B9"].Value2 = Convert.ToString(B9);
                WorkSheet.Range["B10"].Value2 = Convert.ToString(B10);
                WorkSheet.Range["B13"].Value2 = Convert.ToString(B13);
                WorkSheet.Range["B14"].Value2 = Convert.ToString(B14);
                WorkSheet.Range["B15"].Value2 = Convert.ToString(B15);
                WorkSheet.Range["B17"].Value2 = Convert.ToString(B17);
                WorkSheet.Range["B18"].Value2 = Convert.ToString(B18);
                WorkSheet.Range["B22"].Value2 = Convert.ToString(B22);
                WorkSheet.Range["B23"].Value2 = Convert.ToString(B23);



                objExcel.GetType().InvokeMember("Run", BindingFlags.Default | BindingFlags.InvokeMethod,
                    null, objExcel, new Object[] { "Test" });
                B19 = Math.Round(Convert.ToDouble(WorkSheet.Range["B19"].Value), 3);
                B22 = Math.Round(Convert.ToDouble(WorkSheet.Range["B22"].Value), 3);
                B23 = Math.Round(Convert.ToDouble(WorkSheet.Range["B23"].Value), 3);
                B20 = Math.Round(Convert.ToDouble(WorkSheet.Range["B20"].Value), 3);
                B24 = Math.Round(Convert.ToDouble(WorkSheet.Range["B24"].Value), 3);
            }
            catch (Exception ex)
            {
            }
            finally
            {
                if (WorkBook != null) WorkBook.Close(false, null, null);
                if (objExcel != null) objExcel.Quit();
            }
        }

        public ClassRas Rachet()
        {
            return new ClassRas
            {
                B19 = (double)B19,
                B20 = (double)B20,
                B21 = (double)B21,
                B22 = (double)B22,
                B23 = (double)B23,
                B24 = (double)B24
            };
         }
     }
}
