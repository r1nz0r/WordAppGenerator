using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace AppendixGenConsole
{
    static class InputData
    {
        public static List<string> FileNames { get; set; }
        public static List<string> CaptionNames { get; set; }
        public static List<string[]> ResultTypes { get; set; }
        public static List<string[]> ConstructionTypes { get; set; }
        public static List<string[]> CombinationTypes { get; set; }
        public static List<string[]> SubConstructions { get; set; }

        public enum ExcelColumnMeaning
        {
            Name,
            LangEN,
            LangRU,
            None
        }

        public static ExcelColumnMeaning language = ExcelColumnMeaning.LangRU;

        public static void SetInputData (string baseName)
        {
            Dictionary<string, string> captions = new Dictionary<string, string>();
            FileManager.ReadCSV(FileManager.GlobalPath + baseName, ref captions);

            FileNames = new List<string>();
            CaptionNames = new List<string>();

            foreach (var fileName in captions.Keys)
                FileNames.Add(fileName);

            foreach (var fileCaption in captions.Values)
                CaptionNames.Add(fileCaption);
        }

        public static void SetInputData(string excelFilePath, string excelFileName)
        {
            FileNames = FileManager.GetFileNames(false);
            FileNames = HelpClass.SortListByElevations(FileNames);

            Console.Clear();
            Console.WriteLine("Процесс считывания данных из Excel файла займет какое-то время...");

            Excel.Application excelApp = new Excel.Application();
         
            try
            {
                var excelWorkbooks = excelApp.Workbooks;
                var excelWorkbook = excelWorkbooks.Open(excelFilePath + excelFileName);
       
                ResultTypes = FileManager.ReadDataFromExcel(excelWorkbook, FileManager.ExcelWorksheetList.ResultType);
                CombinationTypes = FileManager.ReadDataFromExcel(excelWorkbook, FileManager.ExcelWorksheetList.CombinationType);
                ConstructionTypes = FileManager.ReadDataFromExcel(excelWorkbook, FileManager.ExcelWorksheetList.ConstructionType);
                SubConstructions = FileManager.ReadDataFromExcel(excelWorkbook, FileManager.ExcelWorksheetList.SubStructure);

                if(excelWorkbooks != null)
                {               
                    Marshal.FinalReleaseComObject(excelWorkbooks);
                    excelWorkbooks = null;
                }

                if (excelWorkbook != null)
                {
                    excelWorkbook.Close(true);
                    Marshal.FinalReleaseComObject(excelWorkbook);
                    excelWorkbook = null;
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.FinalReleaseComObject(excelApp);
                    excelApp = null;
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex);
                excelApp.Quit();
                Marshal.FinalReleaseComObject(excelApp);
                excelApp = null;
            }          
        }

        public static ExcelColumnMeaning GetLanguage()
        {
            Console.Clear();
            Console.WriteLine("Выберите язык для вывода названий: ");
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("1 - En");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("2 - Ru");
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.Write("Введите цифрой нужный вариант: ");

            int userInput = 0;

            if (Int32.TryParse(Console.ReadLine(), out userInput) && userInput < (int)ExcelColumnMeaning.None && userInput >= 0)
            {
                return (ExcelColumnMeaning) userInput;
            }

            else return GetLanguage();

        }
    }

}
