using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace TimetableUniter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string docsTimetableData = "";
        private List<string> assistantsTimetablesDataList = new List<string>();

        private string pathToDocTimetable = "";

        private DoctorsTimetableRetriever docRetriever = new DoctorsTimetableRetriever();
        private AssistantTimetableRetriever assistantRetriever = new AssistantTimetableRetriever();
        private TimetablesUniter uniter = new TimetablesUniter();

        public MainWindow()
        {
            InitializeComponent();   
        }



        private void ChooseDocFile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.Filter = "Расписание докторов (*.xlsx) | *.xlsx";

                bool? result = dlg.ShowDialog();

                pathToDocTimetable = dlg.FileName;

                if (result == true)
                    docsTimetableData = docRetriever.RetrieveDoctorsTimetableInformation(pathToDocTimetable, Message);

                Message.Foreground = Brushes.Black;
                Message.Text = "Расписание врачей добавлено.";

                /*
                //var docFileName = @"C:\Users\Daniel3\Desktop\TimetableUniter\TestResources\DoctorsTimetableExample.xlsx";
                var assistantFileName = @"C:\Users\Daniel3\Desktop\TimetableUniter\TestResources\Assistant'sTimetableExample.xlsx";
                var assistantFileName2 = @"C:\Users\Daniel3\Desktop\TimetableUniter\TestResources\Assistant'sTimetableExample2.xlsx";

                // TODO: Make for loop foreach file
                var assistantsTimetableDataList = new List<string>();
                assistantsTimetableDataList.Add(assistantRetriever.RetrieveAssistantsTimetableInformation(assistantFileName, Message));
                assistantsTimetableDataList.Add(assistantRetriever.RetrieveAssistantsTimetableInformation(assistantFileName2, Message));
                */

            }
            catch (Exception ex)
            {
                Message.Foreground = Brushes.Red;
                Message.Text = ex.Message;
            }
        }

        private void ChooseAssistantFile_Click(object sender, RoutedEventArgs e)
        {
            assistantsTimetablesDataList.Clear();

            try
            {
                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.Filter = "Расписания ассистентов (*.xlsx) | *.xlsx";
                dlg.Multiselect = true;

                bool? result = dlg.ShowDialog();

                if (result == true)
                {
                    // Read the files
                    foreach (String file in dlg.FileNames)
                    {
                        assistantsTimetablesDataList.Add(
                            assistantRetriever.RetrieveAssistantsTimetableInformation(file, Message));
                    }
                }

                Message.Foreground = Brushes.Black;
                Message.Text = "Расписание ассистентов добавлено.";
            }
            catch (Exception ex)
            {
                Message.Foreground = Brushes.Red;
                Message.Text = ex.Message;
            }
        }

        private void UniteTimetables_Click(object sender, RoutedEventArgs e)
        {
            bool success = false;

            try
            {
                success = uniter.UniteTimetables(pathToDocTimetable, docsTimetableData, assistantsTimetablesDataList, Message);
            }
            catch (Exception ex)
            {
                Message.Foreground = Brushes.Red;
                Message.Text = ex.Message;
            }

            if (success)
            {
                Message.Foreground = Brushes.Black;
                Message.Text = "Общее расписание создано.";
            }
        }
    }
}
