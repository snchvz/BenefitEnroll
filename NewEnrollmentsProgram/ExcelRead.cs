using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using EXL = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace NewEnrollmentsProgram
{
    public class ExcelRead
    {
        string path = "";

        int count = 0;

        //open 2 excel apps
        _Application excel = new Application();
        _Application excelDest = new Application();

        //used to reference excel process
        int id1 = 0;
        int id2 = 0;

        public ExcelRead(string path, int sheet)
        {

        }

        private string CloseDoc(Workbook wb, Workbook wbDest, Workbooks wbs, Workbooks workbooksDest, dynamic wsDest)
        {
            wbDest.Save();

            wb.Close();
            wbs.Close();

            wbDest.Close();
            workbooksDest.Close();

            //**Excel Process is NOT being terminated after marshall release
            //Marshal.ReleaseComObject(sheets);
            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(wbs);
            Marshal.ReleaseComObject(excel);

            Marshal.ReleaseComObject(wbDest);
            Marshal.ReleaseComObject(workbooksDest);
            Marshal.ReleaseComObject(wsDest);
            Marshal.ReleaseComObject(excelDest);

            try
            {
                Process.GetProcessById(id1).Kill();
                Process.GetProcessById(id2).Kill();
            }
            catch
            {
                return "failed to terminate excel procress\nplease terminate excel process manually through task manager";
            }
            return "Done";
        }

        public string readWriteCell(int monthInt, string year, string path, int sheet, string filename, string destFilename)
        {
            string enrollMonth = monthInt.ToString();

            this.path = path;

            var wbs = excel.Workbooks;
            var wb = wbs.Open(path);
            //var sheets = wb.Worksheets;
            var ws = excel.Sheets[1];            

            var workbooksDest = excelDest.Workbooks;
            var wbDest = workbooksDest.Open(destFilename);
            var wsDest = excelDest.Worksheets[1];

            //**need better way to terminate procress
            try
            {
                var processes = Process.GetProcessesByName("excel");
                id1 = processes[0].Id;
                id2 = processes[1].Id;
            }
            catch
            {                
                return CloseDoc(wb, wbDest, wbs, workbooksDest, wsDest);
            }                      
            
            string month;

            switch (monthInt)
            {
                case 1:
                    monthInt = 10;
                    month = monthInt.ToString();
                    break;
                case 2:
                    monthInt = 11;
                    month = monthInt.ToString();
                    break;
                case 3:
                    monthInt = 12;
                    month = monthInt.ToString();
                    break;
                default:
                    monthInt -= 3;
                    month = monthInt.ToString();
                    break;
            }

            try
            {
                foreach (dynamic worksheet in wbDest.Worksheets)
                {
                    worksheet.Cells.Clear();
                }

                wbDest.Save();
            }
            finally
            {

            }

            int hireCol = 15;
            int termCol = 18;
            int reHireCol = 16;
            int EEIDCol = 1;
            int fNameCol = 2;
            int lNameCol = 3;
            int deptCol = 14;
            int deptPosCol = 20;
            int posCol = 21;

            int i = 2;

            wsDest.Cells[1, 1].Value = "ID";
            wsDest.Cells[1, 2].Value = "Yard";
            wsDest.Cells[1, 3].Value = "LName";
            wsDest.Cells[1, 4].Value = "FName";
            wsDest.Cells[1, 5].Value = "Hire";
            wsDest.Cells[1, 6].Value = "Rehire";
            wsDest.Cells[1, 7].Value = "Pos";
            wsDest.Cells[1, 8].Value = "DeptPos";
            wsDest.Cells[1, 9].Value = "Date";
            wsDest.Cells[1, 10].Value = "EnrollDate";
            wsDest.Cells[1, 11].Value = "LastFriday";

            for (int row = 2; ws.Cells[row, EEIDCol].Value != null; row++)
            {                
                if ( ws.Cells[row, hireCol].Value.GetType() == typeof(DateTime) )
                {
                    if (ws.Cells[row, reHireCol].Value == null)
                    {
                        if ((ws.Cells[row, hireCol].Value.Month.ToString() == month) && ws.Cells[row, hireCol].Value.Year.ToString() == year)
                        {
                            if (ws.Cells[row, termCol].Value != null)
                            {
                                if (ws.Cells[row, termCol].Value.ToString() != "T")
                                {
                                    wsDest.Cells[i, 1].Value = ws.Cells[row, EEIDCol].Value;
                                    wsDest.Cells[i, 2].Value = ws.Cells[row, deptCol].Value;
                                    wsDest.Cells[i, 3].Value = ws.Cells[row, fNameCol].Value;
                                    wsDest.Cells[i, 4].Value = ws.Cells[row, lNameCol].Value;
                                    wsDest.Cells[i, 5].Value = ws.Cells[row, hireCol].Value;
                                    wsDest.Cells[i, 6].Value = ws.Cells[row, reHireCol].Value;
                                    wsDest.Cells[i, 7].Value = ws.Cells[row, posCol].Value;
                                    wsDest.Cells[i, 8].Value = ws.Cells[row, deptPosCol].Value;
                                    wsDest.Cells[i, 9].Value = DateTime.Today.ToShortDateString();
                                    wsDest.Cells[i, 10].Value = enrollMonth + "/01/" + year;
                                    i++;
                                }
                            }
                            else
                            {
                                wsDest.Cells[i, 1].Value = ws.Cells[row, EEIDCol].Value;
                                wsDest.Cells[i, 2].Value = ws.Cells[row, deptCol].Value;
                                wsDest.Cells[i, 3].Value = ws.Cells[row, fNameCol].Value;
                                wsDest.Cells[i, 4].Value = ws.Cells[row, lNameCol].Value;
                                wsDest.Cells[i, 5].Value = ws.Cells[row, hireCol].Value;
                                wsDest.Cells[i, 6].Value = ws.Cells[row, reHireCol].Value;
                                wsDest.Cells[i, 7].Value = ws.Cells[row, posCol].Value;
                                wsDest.Cells[i, 8].Value = ws.Cells[row, deptPosCol].Value;
                                wsDest.Cells[i, 9].Value = DateTime.Today.ToShortDateString();
                                wsDest.Cells[i, 10].Value = enrollMonth + "/01/" + year;
                                i++;
                            }

                        }
                    }
                    else
                    {
                        if ((ws.Cells[row, reHireCol].Value.Month.ToString() == month) && ws.Cells[row, reHireCol].Value.Year.ToString() == year)
                        {
                            if (ws.Cells[row, termCol].Value != null)
                            {
                                if (ws.Cells[row, termCol].Value.ToString() != "T")
                                {
                                    wsDest.Cells[i, 1].Value = ws.Cells[row, EEIDCol].Value;
                                    wsDest.Cells[i, 2].Value = ws.Cells[row, deptCol].Value;
                                    wsDest.Cells[i, 3].Value = ws.Cells[row, fNameCol].Value;
                                    wsDest.Cells[i, 4].Value = ws.Cells[row, lNameCol].Value;
                                    wsDest.Cells[i, 5].Value = ws.Cells[row, hireCol].Value;
                                    wsDest.Cells[i, 6].Value = ws.Cells[row, reHireCol].Value;
                                    wsDest.Cells[i, 7].Value = ws.Cells[row, posCol].Value;
                                    wsDest.Cells[i, 8].Value = ws.Cells[row, deptPosCol].Value;
                                    wsDest.Cells[i, 9].Value = DateTime.Today.ToShortDateString();
                                    wsDest.Cells[i, 10].Value = enrollMonth + "/01/" + year;
                                    i++;
                                }
                            }
                            else
                            {
                                wsDest.Cells[i, 1].Value = ws.Cells[row, EEIDCol].Value;
                                wsDest.Cells[i, 2].Value = ws.Cells[row, deptCol].Value;
                                wsDest.Cells[i, 3].Value = ws.Cells[row, fNameCol].Value;
                                wsDest.Cells[i, 4].Value = ws.Cells[row, lNameCol].Value;
                                wsDest.Cells[i, 5].Value = ws.Cells[row, hireCol].Value;
                                wsDest.Cells[i, 6].Value = ws.Cells[row, reHireCol].Value;
                                wsDest.Cells[i, 7].Value = ws.Cells[row, posCol].Value;
                                wsDest.Cells[i, 8].Value = ws.Cells[row, deptPosCol].Value;
                                wsDest.Cells[i, 9].Value = DateTime.Today.ToShortDateString();
                                wsDest.Cells[i, 10].Value = enrollMonth + "/01/" + year;
                                i++;
                            }

                        }
                    }
                }
            }
            
            return CloseDoc(wb, wbDest, wbs, workbooksDest, wsDest);
        }

        private void KillExcelProcesses()
        {
            var processes = Process.GetProcessesByName("excel");

            foreach (Process p in processes)
                p.Kill();

            //foreach (var process in processes)
            //{
            //    string p = process.;
            //    if (process.proc == excelFileName + " - Microsoft Excel" )
            //        process.Kill();
            //}
        }
    }
}
