using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace Efficiency_Framework.Models
{
    public class InputExcel
    {
        #region Входные данные
        public double[] Rate_PG_Base { get; set; } //Расход природного газа в базовом периоде, м3/ч

        public double[] Rate_PG_Min = new double[6] { 10000, 10000, 10000, 10000, 10000, 10000 }; //Минимально допустимый расход природного газа, м3/ч
        public double[] Rate_PG_Max = new double[6] { 20000, 20000, 20000, 20000, 20000, 20000 }; //Максимально допустимый расход природного газа, м3/ч
        public double[] Rate_Koks_Base { get; set; } //Расход кокса в базовом периоде, т/час
        public double[] Equiv_Replacement { get; set; } //Эквивалент замены кокса в базовом периоде, кг кокса /(м3 ПГ)
        public double[] Performance { get; set; } // Производительность по чугуну в базовом периоде, т /ч
        public double[] Content_S { get; set; } //Содержание S в чугуне в базовом периоде, %
        public double[] Min_S = new double[6] { 0, 0, 0, 0, 0, 0 }; //Минимально допустимое [S], %

        public double[] Max_S = new double[6] { 0.025, 0.025, 0.025, 0.025, 0.025, 0.025 };  //Максимально допустимое [S], %

        public double[] Performance_Change_PG = new double[6] { -0.0007295, -0.0006695, 0, -0.00072373, -0.0007724, -0.0006872 }; //Изменение производства чугуна при изменении ПГ, т чуг/(м3 ПГ)

        public double[] Performance_Change_Koks = new double[6] { -0.00297, -0.00297, -0.002928, -0.002897, -0.00297, -0.00297 }; //Изменение производства чугуна при увеличении расхода кокса, т чуг/(кг кокса)

        public double[] Change_S_PG = new double[6] { -0.0000034, -0.0000034, -0.0000035, -0.0000033, -0.0000034, -0.0000034 }; //Изменение [S] при увеличении расхода ПГ на 1 м3/ч
        public double[] Change_S_Koks = new double[6] { -0.000003, -0.0000029, -0.0000032, -0.0000029, -0.0000031, -0.0000028 }; //Изменение [S] при увеличении расхода кокса на 1 кг/ч

        public double[] Change_S_Performance = new double[6] { 0, 0, 0, 0, 0, 0 }; //Изменение [S] при увеличении производительности печи на 1 т чуг/ч

        public double[] Rate_PG = new double[6] { 0, 0, 0, 0, 0, 0 }; //Расход природного газа, м3/ч
        public double Koks_Cost { get; set; } //Стоимость кокса, руб/(кг кокса)
        public double PG_Cost { get; set; } //Стоимость природного газа, руб/(м3 ПГ)
        public double Reserv_PG { get; set; } //Резерв по расходу природного газа в целом по цеху, м3/ч 
        public double Reserv_Koks { get; set; } //Запасы кокса по цеху, т/ч
        public double Performance_Required { get; set; } //Требуемое производство чугуна в цехе, т/ ч
        #endregion

        Excel.Application objExcel = null;
        Excel.Workbook WorkBook = null;

        public void Solve()
        {
            try
            {
                objExcel = new Excel.Application();

                //objExcel.ScreenUpdating = true;
                //objExcel.WindowState = Excel.XlWindowState.xlMaximized;
                //objExcel.Visible = true;
                //objExcel.DisplayAlerts = true;

                string fileName = Path.Combine(AppContext.BaseDirectory, "Solver.xlsm");

                WorkBook = objExcel.Workbooks.Open(fileName, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                Excel.Worksheet worksheet = (Excel.Worksheet)WorkBook.Sheets["Data"];

                string[] Letters = new string[6] { "D", "E", "F", "G", "H", "I" };
                for (int i = 0; i < 6; i++)
                {
                    worksheet.Range[Letters[i] + "5"].Value2 = Rate_PG_Base[i];
                    worksheet.Range[Letters[i] + "6"].Value2 = Rate_PG_Min[i];
                    worksheet.Range[Letters[i] + "7"].Value2 = Rate_PG_Max[i];
                    worksheet.Range[Letters[i] + "8"].Value2 = Rate_Koks_Base[i];
                    worksheet.Range[Letters[i] + "9"].Value2 = Equiv_Replacement[i];
                    worksheet.Range[Letters[i] + "10"].Value2 = Performance[i];
                    worksheet.Range[Letters[i] + "11"].Value2 = Content_S[i];
                    worksheet.Range[Letters[i] + "12"].Value2 = Min_S[i];
                    worksheet.Range[Letters[i] + "13"].Value2 = Max_S[i];
                    worksheet.Range[Letters[i] + "15"].Value2 = Performance_Change_PG[i];
                    worksheet.Range[Letters[i] + "16"].Value2 = Performance_Change_Koks[i];
                    worksheet.Range[Letters[i] + "17"].Value2 = Change_S_PG[i];
                    worksheet.Range[Letters[i] + "18"].Value2 = Change_S_Koks[i];
                    worksheet.Range[Letters[i] + "19"].Value2 = Change_S_Performance[i];
                    worksheet.Range[Letters[i] + "20"].Value2 = Rate_PG[i];
                }
                worksheet.Range["D25"].Value2 = Koks_Cost;
                worksheet.Range["D26"].Value2 = PG_Cost;
                worksheet.Range["D27"].Value2 = Reserv_PG;
                worksheet.Range["D28"].Value2 = Reserv_Koks;
                worksheet.Range["D29"].Value2 = Performance_Required;

                objExcel.GetType().InvokeMember("Run", System.Reflection.BindingFlags.Default | System.Reflection.BindingFlags.InvokeMethod, null, objExcel, new object[] { "Solve" });

                for (int i = 0; i < 6; i++)
                {
                    Rate_PG[i] = Convert.ToDouble(worksheet.Range[Letters[i] + "20"].Value.ToString("0.##"));
                }
            }
            catch (Exception ex)
            {
                var a = ex.ToString();
            }
            finally
            {
                if (WorkBook != null) WorkBook.Close(false, null, null);
                if (objExcel != null) objExcel.Quit();
            }
        }

        public ResultModel Write() //запись результатов расчета в модель
        {
            Solve();
            return new ResultModel { Rate_PG = Rate_PG };
        }
    }
}