using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace AppendixGenConsole
{
    static class FileManager
    {
        public static string CaptionsFileName;
        public enum ExcelWorksheetList
        {
            None,
            SubStructure,
            ConstructionType,
            ResultType,
            CombinationType,
            Null
        }
        public enum PathType
        {
            Result,
            Global,
            Default
        }

        public static string GlobalPath { get; set; }
        public static string ResultsFolder { get; set; }
        public static string DictionaryFileName { get; set; }

        public static string unbreakSpace = "\u00a0";

        public static List<string> GetFileNames (bool cutExtention = false)

        {
            //Если существует указанная папка
            if (Directory.Exists(GlobalPath + ResultsFolder))
            {
                //Собираем все названия и формируем массив строк
                string[] fileNames = Directory.GetFiles(GlobalPath + ResultsFolder);

                for (int index = 0; index < fileNames.Length; index++)
                {
                    //Удаляем последний символ '\'
                    int pos = fileNames[index].LastIndexOf(@"\") + 1;
                    fileNames[index] = fileNames[index].Substring(pos);

                    //При необходимости удаляем расширение файла (.*)
                    if (cutExtention)
                    {
                        fileNames[index] = RemoveExtentionFromName(fileNames[index]);
                    }

                }
                return new List<string>(fileNames);
            }

            //Если папки нет - просим ввести заново путь к папке и вызываем метод заново
            else
            {
                ResultsFolder = SetResultFolder();
                return GetFileNames(cutExtention);
            }
        }

        public static List<string> GetFileNames (string _path, PathType _pathType)

        {
            //Если существует указанная папка
            if (Directory.Exists(_path))
            {
                string[] names = Directory.GetFiles(_path);

                if (names.Count() == 0)
                {
                    switch (_pathType)
                    {
                        case PathType.Result:
                            return GetFileNames(SetResultFolder(), _pathType);
                        case PathType.Global:
                            return GetFileNames(SetGlobalPath(), _pathType);
                    }
                }

                for (int index = 0; index < names.Length; index++)
                {
                    int pos = names[index].LastIndexOf(@"\") + 1;
                    names[index] = names[index].Substring(pos);
                }

                List<string> listNames = new List<string>(names);
                HelpClass.SortListByElevations(listNames);
                return listNames;
            }

            //Если папки нет - просим ввести заново путь к папке и вызываем метод заново           
            {
                ResultsFolder = SetResultFolder();
                return GetFileNames(_path, _pathType);
            }

        }

        public static string RemoveExtentionFromName (string fileName)
        {
            return fileName.Remove(fileName.LastIndexOf("."));
        }

        public static void RemoveExtentionFromName (ref List<string> _listOfNames)
        {
            for (int index = 0; index < _listOfNames.Count; index++)
            {
                _listOfNames[index] = _listOfNames[index].Remove(_listOfNames[index].LastIndexOf("."));
            }
        }

        public static string SetResultFolder ()
        {
            Console.Clear();
            Console.Write(@"Укажите название локальной папки с результатами нажав Enter: ");

            ResultsFolder = Console.ReadLine();
            ResultsFolder += "\\";

            if (!Directory.Exists(GlobalPath + ResultsFolder) || ResultsFolder == "\\") return SetResultFolder();

            return ResultsFolder;
        }

        public static string SetResultFolder (bool checkPath)
        {
            Console.Clear();
            Console.Write(@"Укажите название локальной папки с результатами нажав Enter: ");
            ResultsFolder = Console.ReadLine();

            if (checkPath)
            {
                ResultsFolder += "\\";
                if (!Directory.Exists(GlobalPath + ResultsFolder) || ResultsFolder == "\\") return SetResultFolder();
            }

            return ResultsFolder;
        }

        public static string SetGlobalPath ()
        {
            Console.Clear();
            Console.Write(@"Введите путь к папке со всеми файлами (");
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.Write(@"E:\Path1\Path2");
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.Write("): ");

            string path = Console.ReadLine() + "\\";

            if (Directory.Exists(path) && path != "\\")
            {
                GlobalPath = path;
                return path;
            }

            else
            {
                return SetGlobalPath();
            }

        }

        public static string SetDictionaryFileName ()
        {
            string defaultFileName = @"dictionary.xlsx";
            Console.Clear();
            Console.Write(@"Укажите название файла со словами-ключами (");
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.Write($"по умолчанию '{defaultFileName}'");
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.Write("), либо пропустите шаг,\nдля использования значения по умолчанию, нажав Enter: ");
            string fileName = Console.ReadLine();
            fileName = (fileName != "") ? fileName : defaultFileName;

            if (!File.Exists(GlobalPath + fileName)) return SetDictionaryFileName();
            return fileName;
        }

        public static List<string[]> ReadCSVFromTXT (string fileName, string globalPath)
        {
            Console.Clear();

            List<string[]> csvData = new List<string[]>();

            try
            {
                var csvReader = new StreamReader(File.OpenRead(globalPath + fileName));
                // Пока не конец документа
                while (!csvReader.EndOfStream)
                {
                    var lineText = csvReader.ReadLine();

                    try
                    {
                        string[] data = lineText.Split(';');
                        csvData.Add(data);
                    }

                    catch
                    {
                        Console.WriteLine("Неверный разделитель в исходном файле (д.б. ';')");
                        break;
                    }
                }
                csvReader.Close();
                return csvData;
            }

            catch
            {
                Console.WriteLine($"Указанный файл {globalPath + fileName} занят другой программой или не существует.");
                return ReadCSVFromTXT(SetDictionaryFileName(), globalPath);
            }

        }

        public static void ReadCSV (string _path, ref Dictionary<string, string> _captionsForPics)

        {
            var csvReader = new StreamReader(File.OpenRead(_path));

            while (!csvReader.EndOfStream)
            {
                var lineText = csvReader.ReadLine();

                try
                {
                    var keyValuePair = lineText.Split(';');
                    _captionsForPics[keyValuePair[0]] = keyValuePair[1];
                }
                catch
                {
                    Console.WriteLine("Неверный разделитель в исходном файле (д.б. ';')");
                    break;
                }
            }
            csvReader.Close();
        }

        public static List<string[]> ReadDataFromExcel (string fileName, string filePath, ExcelWorksheetList ewList)
        {
            List<string[]> data = new List<string[]>();
            Excel.Application exelApp = new Excel.Application();

            try
            {
                Excel.Workbook excelWorkbook = exelApp.Workbooks.Open(filePath + fileName);
                Excel.Worksheet excelWorksheet = excelWorkbook.Worksheets[(int)ewList];

                int row = 2;
                int column = 1;
                string valueOfCell = "";

                while (excelWorksheet.Cells[row, 1].Value2 != null)
                {
                    string[] excelLineData = new string[3];

                    for (column = 1; column <= 3; column++)
                    {
                        valueOfCell = excelWorksheet.Cells[row, column].Value2;
                        excelLineData[column - 1] = valueOfCell.Replace("{unbrS}", unbreakSpace);
                    }

                    data.Add(excelLineData);
                    row++;
                }

                excelWorkbook.Close();
                exelApp.Quit();
                return data;

            }

            catch (Exception ex)
            {
                Console.WriteLine(ex);
                exelApp.Quit();
            }

            return null;
        }

        public static List<string[]> ReadDataFromExcel (Excel.Workbook excelWorkbook, ExcelWorksheetList ewList)
        {

            List<string[]> data = new List<string[]>();

            try
            {
                Excel.Worksheet excelWorksheet = excelWorkbook.Worksheets[(int)ewList];

                int row = 2;
                int column = 1;
                string valueOfCell = "";


                var cellsForCheck = excelWorksheet.Cells[row, 1];
                var valueForCheck = cellsForCheck.Value2;

                while (valueForCheck != null)
                {
                    string[] excelLineData = new string[3];

                    for (column = 1; column <= 3; column++)
                    {
                        cellsForCheck = excelWorksheet.Cells[row, column];
                        valueOfCell = cellsForCheck.Value2;
                        excelLineData[column - 1] = valueOfCell.Replace("{unbrS}", unbreakSpace);
                    }

                    data.Add(excelLineData);
                    row++;

                    cellsForCheck = excelWorksheet.Cells[row, 1];
                    valueForCheck = cellsForCheck.Value2;
                }

                return data;
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

            return null;
        }

        public static List<string> GetAccelerationList (List<string> fileNames, AppendixFRS.Damping damp)
        {
            List<string> accelerationLast = new List<string>();

            foreach (var s in fileNames)
            {
                List<string> accelerationsAll = new List<string>();

                var csvReader = new StreamReader(File.OpenRead(FileManager.ResultsFolder + s));
                while (!csvReader.EndOfStream)
                {
                    var lineText = csvReader.ReadLine();
                    try
                    {
                        var accelPerDamp = lineText.Split(';');
                        double accelNum;
                        Double.TryParse(accelPerDamp[(int)damp], out accelNum);
                        string accelStr = accelNum.ToString("0.000") + " m/s^2";
                        accelerationsAll.Add(accelStr);
                    }
                    catch
                    {
                        break;
                    }
                }
                csvReader.Close();
                accelerationLast.Add(accelerationsAll.Last());
            }

            return accelerationLast;
        }

        public static void WriteCSV (string _path, List<PictureData> _pics)
        {
            var stream = new StreamWriter(_path, false, System.Text.Encoding.Unicode);

            foreach (var pic in _pics)
            {
                stream.WriteLine(pic.PictureName + ";" + pic.FullCaption);
            }
            stream.Close();
        }

        public static void WriteCSV (string _path, List<string> names)
        {
            var stream = new StreamWriter(_path, false, System.Text.Encoding.Unicode);

            foreach (var name in names)
            {
                stream.WriteLine(name + ";" + name);
            }
            stream.Close();
        }

        public static void WriteCSV (string _path, Dictionary<string, string> namesKVP)
        {
            var stream = new StreamWriter(_path, false, System.Text.Encoding.Unicode);

            foreach (var name in namesKVP)
            {
                stream.WriteLine(name.Key + ";" + name.Value);
            }
            stream.Close();
        }

        public static void ShowSaveCSVMessage ()
        {
            Console.Clear();
            Console.Write(@"Файл ");
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.Write($"{FileManager.GlobalPath + FileManager.CaptionsFileName}");
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine(" с названиями картинок успешно создан. Необходимо его откорректировать");
            Console.Write("Для продолжения нажмите любую клавишу...");
            Console.ReadKey();
            Console.Clear();
        }
    }
}
