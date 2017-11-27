using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Controls;

namespace TimetableUniter
{
    class DoctorsTimetableRetriever
    {
        private static readonly int excelDataCapacity = 2000;

        private static readonly int StartRow = 2;
        private static readonly int StartColumn = 2;
        private static readonly int FinishRow = 32;
        private static readonly int FinishColumn = 3;

        private string fileName;

        // Create COM Objects. Create a COM object for everything that is referenced.
        Application xlApp;
        Workbook xlWorkbook;
        Worksheet xlWorksheet;
        Range xlRange;

        public string RetrieveDoctorsTimetableInformation(string sourceFile, TextBlock errorMessage)
        {
            // Cache source file name
            fileName = sourceFile;

            try
            {
                InitializeExcelReferences();
                return RetrieveDataFromExcel();
            }
            finally
            {
                CleanExcelReferences();
            }
        }



        private void InitializeExcelReferences()
        {
            xlApp = new Application();
            xlWorkbook = xlApp.Workbooks.Open(fileName);
            xlWorksheet = xlWorkbook.Sheets[1];
            xlRange = xlWorksheet.UsedRange;
        }

        private string RetrieveDataFromExcel()
        {
            var data = new StringBuilder(excelDataCapacity);

            for (int i = StartRow; i <= FinishRow; i++)
            {
                for (int j = StartColumn; j <= FinishColumn; j++)
                {
                    //write the value to data
                    if (xlRange.Cells[i, j] != null &&
                        xlRange.Cells[i, j].Value2 != null)
                    {
                        data.Append(xlRange.Cells[i, j].Value2.ToString() + ";");
                    }
                    else data.Append(";");
                }
            }
            // Removing last ';', because otherwise string.split() adds one excessive element
            data.Remove(data.Length - 1, 1);

            return data.ToString();
        }

        private void CleanExcelReferences()
        {
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            if (xlRange != null) Marshal.ReleaseComObject(xlRange);
            if (xlWorksheet != null) Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            if (xlWorkbook != null)
            {
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
            }

            //quit and release
            if (xlApp != null)
            {
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
        }
    }
}