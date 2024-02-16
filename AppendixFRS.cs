using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AppendixGenConsole
{
    class AppendixFRS
    {     
        public enum Damping
        {
            Freq,
            D_05,
            D_1,
            D_2,
            D_3,
            D_4,
            D_5,
            D_7,
            D_10,
            D_15,
            Null
        }
        public enum FRSType
        {     
            NONE,
            SEISM,
            ASW,
            ACCEL
        }
        public enum Dictionary
        {
            Name,
            En,
            Ru
        }

        static string resultAccelPicsPath;
        static string resultDisplPicsPath;
        static FRSType typeOfFRS;
        public static object oMissing = System.Reflection.Missing.Value;
        public static string globalPath = @"E:\WORK\VS\TEST_FRS\";
        public static Dictionary dictionary = Dictionary.En;
        static WordHelper wordHelper;

        public static void MenuFRS()
        {

            Console.Clear();
            Console.WriteLine("Выберите необходимое действие: ");
            Console.WriteLine("1. Генерация приложения с результатами спектров для сейсмики;");
            Console.WriteLine("2. Генерация приложения с результатами спектров для ВУВ/Самолета;");
            Console.WriteLine("3. Вычленить значения нулевого периода по отметкам.");

            Console.Write("Для этого введите число от 1 до 3 и нажмите Enter: ");
            int action = 0;
            Int32.TryParse(Console.ReadLine(), out action);
            typeOfFRS = (FRSType) action;
            Console.Clear();
            if (typeOfFRS == FRSType.ACCEL)
                GetAllAccelerationsPerFile(true);
            else
                GenerateAppendix();            
        }
        public static void GenerateAppendix()
        {            
            string labelCaptionPartOne = "";
            string[] captionForColumn = new string[10];

            Console.Clear();

            InitializeFileManager(typeOfFRS);
            

            GetPicsPath(out resultAccelPicsPath, out resultDisplPicsPath);

            List<string> accelerationList = GetAccelerationsFromPath(resultAccelPicsPath);

            //Определяемся с языком, на котором будет осуществляться построение описания к картинкам
            Dictionary languageInUse = GetLanguage();

            int fileFormat = GetFileFormat();

            List<string> elevations;
            Dictionary<string, string> resultNameDictFull;

            CollectInitialData(ref labelCaptionPartOne, ref captionForColumn, accelerationList, languageInUse, fileFormat, out elevations, out resultNameDictFull);

            string labelCaptionPartTwo = GetLabelCaptionPartTwo();
            InitializeWordHelper();
            string fullLabelCaption;
            object oLabelCaption;
            GetFullLabelCaption(labelCaptionPartOne, labelCaptionPartTwo, out fullLabelCaption, out oLabelCaption);
            InsertContentInDocument(captionForColumn, accelerationList, languageInUse, fileFormat, elevations, resultNameDictFull, fullLabelCaption, oLabelCaption);

            wordHelper.App.Visible = true;
            Save();
        }

        private static object InsertContentInDocument(string[] captionForColumn, List<string> accelerationList, Dictionary languageInUse, int userInput, List<string> elevations, Dictionary<string, string> resultNameDictFull, string fullLabelCaption, object oLabelCaption)
        {
            if (userInput == 1)
            {
                GenerateAppWithElevationList(captionForColumn, accelerationList, languageInUse, elevations, fullLabelCaption, oLabelCaption);
                CreateTableWithReferences(elevations, ref oLabelCaption);
            }
            else
            {
                if (typeOfFRS == FRSType.ASW)
                {
                    wordHelper.App.Selection.InsertParagraphAfter();
                    wordHelper.App.Selection.InsertParagraphAfter();
                    string pictureCaptionPt2 = $"(Поэтажные спектры отклика и поэтажные спектры перемещений соответственно)";

                    foreach (var result in resultNameDictFull)
                    {
                        string accelPic = resultAccelPicsPath + result.Key;
                        string displPic = resultDisplPicsPath + result.Key;

                        wordHelper.InsertTable(1, 2);
                        wordHelper.FormatTableForFRS(captionForColumn, accelPic, displPic);
                        wordHelper.InsertCaptionForPic(oLabelCaption, fullLabelCaption, result.Value, pictureCaptionPt2);
                    }
                }

                else
                {
                    foreach (var result in resultNameDictFull)
                    {
                        string accelPic = resultAccelPicsPath + result.Key;

                        wordHelper.InsertPicture(accelPic);
                        wordHelper.InsertInfoTextInTable(captionForColumn);
                        wordHelper.InsertCaptionForPic(oLabelCaption, fullLabelCaption, result.Value);
                    }
                }

                CreateTableWithReferences(resultNameDictFull, ref oLabelCaption);
            }

            return oLabelCaption;
        }

        private static void GetPicsPath(out string resultAccelPicsPath, out string resultDisplPicsPath)
        {
            if (typeOfFRS == FRSType.ASW)
            {
                resultAccelPicsPath = FileManager.GlobalPath + FileManager.ResultsFolder + @"_accel_pictures\";
                resultDisplPicsPath = FileManager.GlobalPath + FileManager.ResultsFolder + @"_displ_pictures\";
            }
            else
            {
                resultAccelPicsPath = FileManager.GlobalPath + FileManager.ResultsFolder + @"_pictures\";
                resultDisplPicsPath = "";
            }
        }

        private static void CollectInitialData(ref string labelCaptionPartOne, ref string[] captionForColumn, List<string> accelerationList, Dictionary languageInUse, int fileFormat, out List<string> elevations, out Dictionary<string, string> resultNameDictFull)
        {
            SetLanguage(ref captionForColumn, ref labelCaptionPartOne, languageInUse);

            if (fileFormat == 1)
            {
                elevations = GetElevations(accelerationList, languageInUse);
                resultNameDictFull = null;
            }
            else
            {
                elevations = GetResultAccelerationNames(accelerationList, languageInUse);
                resultNameDictFull = GenerateDataFilesCSV(accelerationList, elevations);
            }
        }

        private static void GenerateAppWithElevationList(string[] captionForColumn, List<string> accelerationList, Dictionary languageInUse, List<string> elevations, string fullLabelCaption, object oLabelCaption)
        {
            if (typeOfFRS == FRSType.ASW)
            {
                wordHelper.App.Selection.InsertParagraphAfter();
                wordHelper.App.Selection.InsertParagraphAfter();
                for (int i = 0; i < accelerationList.Count; i++)
                {
                    string accelPicPath = resultAccelPicsPath + accelerationList[i] + ".png";
                    string displPicPath = resultDisplPicsPath + accelerationList[i] + ".png";
                    string pictureNamePt1 = "";
                    string pictureNamePt2 = "";
                    GetCaptions(elevations[i], accelerationList[i], ref pictureNamePt1, ref pictureNamePt2, languageInUse);
                    wordHelper.InsertTable(1, 2);
                    wordHelper.FormatTableForFRS(captionForColumn, accelPicPath, displPicPath);
                    wordHelper.InsertCaptionForPic(oLabelCaption, fullLabelCaption, pictureNamePt1, pictureNamePt2);
                }

            }

            else if (typeOfFRS == FRSType.SEISM)
            {
                for (int i = 0; i < accelerationList.Count; i++)
                {
                    string accelPicPath = resultAccelPicsPath + accelerationList[i] + ".png";
                    string pictureNamePt1 = "";
                    string pictureNamePt2 = "";
                    GetCaptions(elevations[i], accelerationList[i], ref pictureNamePt1, ref pictureNamePt2, languageInUse);

                    wordHelper.InsertPicture(accelPicPath);
                    wordHelper.InsertInfoTextInTable(captionForColumn);
                    wordHelper.InsertCaptionForPic(oLabelCaption, fullLabelCaption, pictureNamePt1);
                }
            }
        }

        private static void GetFullLabelCaption(string labelCaptionPartOne, string labelCaptionPartTwo, out string fullLabelCaption, out object oLabelCaption)
        {
            fullLabelCaption = labelCaptionPartOne + labelCaptionPartTwo;
            wordHelper.App.CaptionLabels.Add(fullLabelCaption);
            oLabelCaption = wordHelper.App.CaptionLabels[fullLabelCaption];
        }

        private static void InitializeWordHelper()
        {
            wordHelper = new WordHelper();
            wordHelper.OpenApplication();
            wordHelper.CreateDoc();
            wordHelper.App.Visible = false;
            wordHelper.Rng = wordHelper.Doc.Range();
        }

        private static string GetLabelCaptionPartTwo()
        {
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.Write("Введите основное название подписи к рисункам (напр. ");
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.Write("F.");
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.Write("): ");
            string labelCaptionPartTwo = Console.ReadLine();
            return labelCaptionPartTwo;
        }

        private static Dictionary<string, string> GenerateDataFilesCSV(List<string> accelerationList, List<string> elevations)
        {
            FileManager.CaptionsFileName = $"captions_{FileManager.ResultsFolder.Replace("\\", "")}.csv";
            FileManager.WriteCSV(FileManager.GlobalPath + FileManager.CaptionsFileName, elevations);
            FileManager.ShowSaveCSVMessage();

            Dictionary<string, string> resultNameDict = new Dictionary<string, string>();
            FileManager.ReadCSV(FileManager.GlobalPath + FileManager.CaptionsFileName, ref resultNameDict);

            Dictionary<string, string> resultNameDictFull = new Dictionary<string, string>();

            foreach (var kvp in resultNameDict)
            {
                string resultCaptionName = kvp.Value;
                string unbreakSpace = "\u00a0";
                
                string fullAxis = "";

                for (int axisNum = 0; axisNum < 3; axisNum++)
                {
                    string axisKey = "";
                    switch (axisNum)
                    {
                        case 0:
                            {
                                fullAxis = $"Горизонтальная компонента{unbreakSpace}X";
                                axisKey = "x";
                                break;
                            }
                        case 1:
                            {
                                fullAxis = $"Горизонтальная компонента{unbreakSpace}Y";
                                axisKey = "y";
                                break;
                            }
                        case 2:
                            {
                                fullAxis = $"Вертикальная компонента{unbreakSpace}Z";
                                axisKey = "z";
                                break;
                            }
                        default:
                            {
                                fullAxis = "Нет информации";
                                break;
                            }
                    }

                    string pictureNamePt1 = $" – {resultCaptionName}. {fullAxis}";
                    resultNameDictFull.Add(kvp.Key + $".{axisKey}.png", pictureNamePt1);
                    string pictureNamePt2 = $"(Поэтажные спектры отклика и поэтажные спектры перемещений соответственно)";
                }
                
            }

            //for (int i = 0; i < accelerationList.Count; i++)
            //{
            //    int posOfPoint = accelerationList[i].LastIndexOf('.');
            //    string subString = accelerationList[i];
            //    subString = subString.Substring(0, posOfPoint);

            //    string resultCaptionName = resultNameDict[subString];

            //    string unbreakSpace = "\u00a0";
            //    char axis = accelerationList[i].ElementAt<char>(accelerationList[i].Length - 1);
            //    string fullAxis = "";

            //    switch (axis)
            //    {
            //        case 'x':
            //            {
            //                fullAxis = $"Горизонтальная компонента{unbreakSpace}X";
            //                break;
            //            }
            //        case 'y':
            //            {
            //                fullAxis = $"Горизонтальная компонента{unbreakSpace}Y";
            //                break;
            //            }
            //        case 'z':
            //            {
            //                fullAxis = $"Вертикальная компонента{unbreakSpace}Z";
            //                break;
            //            }
            //        default:
            //            {
            //                fullAxis = "Нет информации";
            //                break;
            //            }
            //    }
            //    string pictureNamePt1 = $" – {resultCaptionName}. {fullAxis}";

            //    resultNameDictFull.Add(accelerationList[i] + ".png", pictureNamePt1);

            //    string pictureNamePt2 = $"(Поэтажные спектры отклика и поэтажные спектры перемещений соответственно)";
            //}

            FileManager.CaptionsFileName = $"captions_{FileManager.ResultsFolder.Replace("\\", "")}_full.csv";
            FileManager.WriteCSV(FileManager.GlobalPath + FileManager.CaptionsFileName, resultNameDictFull);
            FileManager.ShowSaveCSVMessage();

            FileManager.ReadCSV(FileManager.GlobalPath + FileManager.CaptionsFileName, ref resultNameDictFull);

            return resultNameDictFull;
        }

        private static int GetFileFormat()
        {
            Console.Clear();
            Console.WriteLine("Выберите ваш формат названия картинок:");
            Console.WriteLine("1 - level_+5.000.x.png");
            Console.WriteLine("2 - Steamgenerator.x.png");
            int userInput = Int32.Parse((Console.ReadLine()));
            return userInput;
        }

        public static List<string> GetElevations(List<string> fileNames, Dictionary languageInUse)
        {
            List<string> elevations = new List<string>();

            for (int index = 0; index < fileNames.Count; index++)
            {
                int posOfPoint = fileNames[index].LastIndexOf('.');
                int posOfUndescore = fileNames[index].LastIndexOf('_');
                int lengthOfString = posOfPoint - posOfUndescore;
                string subString = fileNames[index];
                subString = subString.Substring(posOfUndescore + 1, lengthOfString - 1);
                if (languageInUse == Dictionary.Ru) subString = subString.Replace("-", "минус\u00a0");
                else subString = subString.Replace("-", "minus\u00a0");
                elevations.Add(subString);
            }

            return elevations;
        }

        public static void InitializeFileManager(FRSType typeOfFRS)
        {
            Console.Clear();            
            FileManager.GlobalPath = FileManager.SetGlobalPath();           
            FileManager.ResultsFolder = FileManager.SetResultFolder(false);
        }
        static Dictionary GetLanguage()
        {
            Console.WriteLine("Выберите язык для вывода названий: ");
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("1 - En");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("2 - Ru");
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.Write("Введите цифрой нужный вариант: ");
            return (Dictionary)int.Parse(Console.ReadLine());
        }
        static void SetLanguage(ref string[] captionForColumn, ref string labelCaption, Dictionary language)
        {
            if (typeOfFRS == FRSType.ASW)
            {
                if (language == Dictionary.Ru)
                {
                    string[] _captions = {"Кривые соответствуют относительным затуханиям:",
                    //"- 0.5 % (верхняя кривая);",
                    "- 1 %;",
                    "- 2 %;",
                    "- 3 %;",
                    "- 4 %;",
                    "- 5 %;",
                    "- 7 %;",
                    "- 10 %;",
                    "- 15 % (нижняя кривая)." };
                    captionForColumn = _captions;
                    labelCaption = "Рисунок ";
                }

                else
                {
                    string[] _captions = {"The curves correspond to relative damping values:",
                    //"— 0.5 % (upper curve);",
                    "— 1 %;",
                    "— 2 %;",
                    "— 3 %;",
                    "— 4 %;",
                    "— 5 %;",
                    "— 7 %;",
                    "— 10 %;",
                    "— 15 % (lower curve)." };
                    captionForColumn = _captions;
                    labelCaption = "Figure ";
                }
            }

            else
            {
                if (language == Dictionary.Ru)
                {
                    string[] _captions = {"Кривые соответствуют относительным затуханиям:",
                    //"- 0.5 % (верхняя кривая);",
                    "- 1 %;",
                    "- 2 %;",
                    "- 3 %;",
                    "- 4 %;",
                    "- 5 %;",
                    "- 7 %;",
                    "- 10 %;",
                    "- 15 % (нижняя кривая)." };
                    captionForColumn = _captions;
                    labelCaption = "Рисунок ";
                }

                else
                {
                    string[] _captions = {"The curves correspond to relative damping values:",
                    //"- 0.5 % (upper curve);",
                    "— 1 %;",
                    "— 2 %;",
                    "— 3 %;",
                    "— 4 %;",
                    "— 5 %;",
                    "— 7 %;",
                    "— 10 %;",
                    "— 15 % (lower curve)." };
                    captionForColumn = _captions;
                    labelCaption = "Figure ";
                }
            }
            
        }
        static void GetCaptions(string elevation, string nameOfPic, ref string pictureNamePt1, ref string pictureNamePt2, Dictionary languageInUse)
        {
            string unbreakSpace = "\u00a0";
            char axis = nameOfPic.ElementAt<char>(nameOfPic.Length - 1);
            //string pictureName = $" — Плита на отметке{unbreakSpace}{level}. Вертикальная компонента{unbreakSpace}Z" +
            //    $"{newLine}(Поэтажные спектры отклика и поэтажные спектры перемещений соответственно)";
            if (languageInUse == Dictionary.Ru)
            {
                string fullAxis = "";
                switch (axis)
                {
                    case 'x':
                        {
                            fullAxis = $"Горизонтальная компонента{unbreakSpace}X";
                            break;
                        }
                    case 'y':
                        {
                            fullAxis = $"Горизонтальная компонента{unbreakSpace}Y";
                            break;
                        }
                    case 'z':
                        {
                            fullAxis = $"Вертикальная компонента{unbreakSpace}Z";
                            break;
                        }
                    default:
                        {
                            fullAxis = "Нет информации";
                            break;
                        }
                }
                pictureNamePt1 = $" – Плита на отметке{unbreakSpace}{elevation}. {fullAxis}";
                pictureNamePt2 = $"(Поэтажные спектры отклика и поэтажные спектры перемещений соответственно)";
            }
            else
            {
                string fullAxis = "";
                switch (axis)
                {
                    case 'x':
                        {
                            fullAxis = $"Horizontal component{unbreakSpace}X";
                            break;
                        }
                    case 'y':
                        {
                            fullAxis = $"Horizontal component{unbreakSpace}Y";
                            break;
                        }
                    case 'z':
                        {
                            fullAxis = $"Vertical component{unbreakSpace}Z";
                            break;
                        }
                    default:
                        {
                            fullAxis = "No info";
                            break;
                        }
                }
                pictureNamePt1 = $" – Slab at elevation{unbreakSpace}{elevation}. {fullAxis}";
                pictureNamePt2 = $"(Floor response spectra and floor displacement spectra respectively)";
            }

        }
        public static void GetAllAccelerationsPerFile(bool sort = true)
        {
            FileManager.SetResultFolder();
            Damping damp = SetDamping();

            List<string> fileNames = FileManager.GetFileNames(FileManager.ResultsFolder, FileManager.PathType.Result);

            if (sort) HelpClass.SortListByElevationsFRS(ref fileNames); // сортировка названий от самой нижней отметки до самой верхней

            List<string> accelerationLast = FileManager.GetAccelerationList(fileNames, damp);
            Dictionary<string, string> accelerations = new Dictionary<string, string>();

            for (int i = 0; i < fileNames.Count; i++)
            {
                accelerations.Add(fileNames[i], accelerationLast[i]);
            }

            Console.Clear();

            int countIteration = 0;
            foreach (var acceleration in accelerations)
            {
                if (countIteration == 3)
                {
                    countIteration = 0;
                    Console.WriteLine();
                }
                countIteration++;
                Console.WriteLine(acceleration.Key + "\t" + acceleration.Value);
            }
        }
        static Damping SetDamping()
        {
            Console.Clear();
            Console.WriteLine("Укажите для какого затухания брать значения:");
            for (int i = 1; i < (int)Damping.Null; i++)
            {
                Console.WriteLine($"{i}. Демпфирование равное {((Damping)i).ToString()}");
            }
            Console.Write($"Введите число от 1 до {(int)Damping.Null - 1}: ");

            int selectedDamp;
            bool correctInput = Int32.TryParse(Console.ReadLine(), out selectedDamp);

            return correctInput ? (Damping)selectedDamp : SetDamping();
        }
        static void CreateTableWithReferences(List<string> elevations, ref object oLabelCaption)
        {
            Console.Clear();
            Console.Write("Желаете создать в начале документа таблицу ссылок на рисунки? (1 - да: 2 - нет): ");

            int userInput = int.Parse(Console.ReadLine());
            if (userInput == 1)
            {
                wordHelper.MoveToBeginningOfTheDoc();
                wordHelper.App.Selection.InsertParagraphBefore();
                wordHelper.MoveToBeginningOfTheDoc();

                int rowsCount = elevations.Count / 3;
                wordHelper.InsertTableFRS(rowsCount, 5);
                wordHelper.FillHeadOfFRSTable();
                wordHelper.FillBodyOfFRSTable(elevations, ref oLabelCaption);
            }
        }

        static void CreateTableWithReferences(Dictionary<string, string> resultDict, ref object oLabelCaption)
        {
            Console.Clear();
            Console.Write("Желаете создать в начале документа таблицу ссылок на рисунки? (1 - да: 2 - нет): ");

            List<string> resultList = new List<string>();

            foreach (var kvp in resultDict)
            {
                resultList.Add(kvp.Value);
            }

            int userInput = int.Parse(Console.ReadLine());
            if (userInput == 1)
            {
                wordHelper.MoveToBeginningOfTheDoc();
                wordHelper.App.Selection.InsertParagraphBefore();
                wordHelper.MoveToBeginningOfTheDoc();

                int rowsCount = resultDict.Count / 3;
                wordHelper.InsertTableFRS(rowsCount, 5);
                wordHelper.FillHeadOfFRSTable();
                wordHelper.FillBodyOfFRSTable(resultList, ref oLabelCaption);
            }
        }

        public static List<string> GetResultAccelerationNames(List<string> fileNames, Dictionary languageInUse)
        {
            List<string> resultNames = new List<string>();

            for (int index = 0; index < fileNames.Count; index+=3)
            {
                int posOfPoint = fileNames[index].LastIndexOf('.');
                string subString = fileNames[index];
                subString = subString.Substring(0, posOfPoint);
                resultNames.Add(subString);
            }

            return resultNames;
        }
        static List<string> GetAccelerationsFromPath(string path)
        {
            List <string> accelerationList = FileManager.GetFileNames(path, FileManager.PathType.Result);
            HelpClass.RemoveExtentionFromName(ref accelerationList);
            HelpClass.SortListByElevationsFRS(ref accelerationList);

            return accelerationList;
        }
        static void Save()
        {
            Console.WriteLine("Документ готов! Для сохранения нажмите любую кнопку.");
            Console.ReadKey(true);
            Console.Write("Сохраняем документ...");

            string name = "Appendix_" + FileManager.ResultsFolder;
            string path = FileManager.GlobalPath + name;

            wordHelper.SaveDoc(path, name);
        }
    }
}
