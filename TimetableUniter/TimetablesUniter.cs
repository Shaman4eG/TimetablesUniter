using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Controls;
using System.Windows.Media;

namespace TimetableUniter
{
    // TODO: Give month choice and remove hardcode
    class TimetablesUniter
    {
        private static readonly int StartRow = 2;
        private static readonly int StartColumn = 2;
        private static readonly int FinishRow = 32;
        private static readonly int FinishColumn = 3;

        private static readonly int MaxShiftsInMonth = 62;
        private static readonly int daysInMonth = 31;

        private static readonly int pairStringCapacity = 200;

        private string month = "Декабрь";
        private DateTime dayOfMonth = new DateTime(2017, 12, 1);


        // Create COM Objects. Create a COM object for everything that is referenced.
        Application xlApp;
        Workbook xlWorkbook;
        Worksheet xlWorksheet;
        Range xlRange;

        public bool UniteTimetables(
            string outputPath,
            string docsTimetable, 
            List<string> assistantsTimetables, 
            TextBlock message)
        {
            try
            {
                var valid = CheckInputFromExcelValidness(docsTimetable, assistantsTimetables);
                if (!valid)
                {
                    message.Foreground = Brushes.Red;
                    message.Text = "Не удалось создать общее расписание. Расписание врачей или расписания ассистентов не были добавлены.";
                    return false;
                }

                InitializeExcelReferences();
                FillFile(docsTimetable, assistantsTimetables);
                SaveFile();

                return true;
            }
            finally
            {
                CleanExcelReferences();
            }
        }



        private void InitializeExcelReferences()
        {
            xlApp = new Application();
            xlWorkbook = xlApp.Workbooks.Add("");
            xlWorksheet = xlWorkbook.ActiveSheet;
        }

        // TODO: TEST AND SEPARATE IN DIFFERENT FUNCTIONS
        private void FillFile(string docsTimetable, List<string> assistantsTimetables)
        {
            // Set alignment for headers
            xlRange = xlWorksheet.get_Range("A1", "C1");
            xlRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // Set alignment for dates
            xlRange = xlWorksheet.get_Range("A2", "A32");
            xlRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // Add borders
            xlRange = xlWorksheet.get_Range("A1", "C32");
            Borders border = xlRange.Borders;
            border.LineStyle = XlLineStyle.xlContinuous;
            border.Weight = 2d;

            // Add table headers.
            xlWorksheet.Cells[1, 1] = month;
            xlWorksheet.Cells[1, 2] = "Утро";
            xlWorksheet.Cells[1, 3] = "Вечер";

            // Add month's dates.
            // "+ 1" cause starting from second row.
            for (int i = 2; i <= daysInMonth + 1; i++)
            {
                xlRange.Cells[i, 1] = "'" + dayOfMonth.ToShortDateString();
                dayOfMonth = dayOfMonth.AddDays(1);
            }

            // Prepare data
            var docsTimetableByDayTime = new List<string>(docsTimetable.Split(';'));

            var assistants = new List<Assistant>();
            foreach (string assistantData in assistantsTimetables)
            {
                string[] assistantTimetableByDayTime = assistantData.Split(';');
                var assistant = new Assistant { LastName = assistantTimetableByDayTime[0] };
                for (int i = 1; i < assistantTimetableByDayTime.Length; i++)
                    if (assistantTimetableByDayTime[i] == "1") assistant.AddShift(i - 1, true);

                assistants.Add(assistant);
            }

            /*
            - Не надо emptyEntries убирать???
            */
            // Fill main data
            var pairs = new List<StringBuilder>();
            for (int shift = 0; shift < MaxShiftsInMonth; shift++)
            {
                pairs.Add(new StringBuilder(pairStringCapacity));

                // Get staff for one shift.
                var assistantsForShift = GetPossibleAssistantsForShift(shift, assistants);
                var docsforShift = new List<string>(docsTimetableByDayTime[shift].Split(
                    new char[] { ',', ' '}, StringSplitOptions.RemoveEmptyEntries));

                // If there are 0 assistants in shift, doesn't output ", " before first doctor's lastname.
                if (assistantsForShift.Count == 0 && docsforShift.Count > 0)
                {
                    pairs[shift].Append(docsforShift[0]);
                    docsforShift.RemoveAt(0);
                }

                // Foreach doc add assistant.
                while (docsforShift.Count != 0 && assistantsForShift.Count != 0)
                {
                    // Choose assistant with lowest number of shifts. 
                    if (docsforShift.Count < assistantsForShift.Count)
                    {
                        // TEST
                        assistantsForShift.Sort((x, y) => x.NumberOfShifts.CompareTo(y.NumberOfShifts));
                    }
                    var chosenAssistant = assistantsForShift[0];
                    chosenAssistant.NumberOfShifts++;

                    // Pair assistant and doc.
                    if (docsforShift.Count != 1 && assistantsForShift.Count != 1)
                        pairs[shift].Append(docsforShift[0] + "-" + chosenAssistant.LastName + ", ");
                    else pairs[shift].Append(docsforShift[0] + "-" + chosenAssistant.LastName);

                    // Remove paired doc and assistent from order to get paired in this shift.
                    docsforShift.RemoveAt(0);
                    assistantsForShift.RemoveAt(0);
                }


                // Add doc to output, if he is left without pair.
                foreach (var doc in docsforShift)
                    pairs[shift].Append(", " + doc);
            }

            // Add pairs to file.
            for (int i = StartRow; i <= FinishRow; i++)
            {
                for (int j = StartColumn; j <= FinishColumn; j++)
                {
                    var pair = pairs[0].ToString();
                    if (pair != "") xlWorksheet.Cells[i, j] = pair;
                    pairs.RemoveAt(0);
                }
            }

            // Fit width
            xlRange = xlWorksheet.get_Range("A1", "C1");
            xlRange.EntireColumn.AutoFit();
        }

        private List<Assistant> GetPossibleAssistantsForShift(int shiftIndex, List<Assistant> assistants)
        {
            var possibleAssistants = new List<Assistant>();

            foreach(var assistant in assistants)
            {
                if (assistant.GetShift(shiftIndex) == true)
                    possibleAssistants.Add(assistant);
            }

            return possibleAssistants;
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

        private bool CheckInputFromExcelValidness(string docsTimetable, List<string> assistantsTimetables)
        {
            if (docsTimetable == "") return false;
            if (assistantsTimetables.Count == 0) return false;

            return true;
        }

        private void SaveFile()
        {
            var desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var fullFileName = Path.Combine(desktopFolder, "Объединенное расписание.xlsx");

            // var outputPath = @"C:\Users\Daniel3\Desktop\TimetableUniter\UnitedTable\UnitedTable.xlsx";
            xlWorkbook.SaveAs(fullFileName, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }
    }
}
 