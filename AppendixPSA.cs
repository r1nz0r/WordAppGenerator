using System;
using System.Collections.Generic;
using System.Linq;

namespace AppendixGenConsole
{
    static class AppendixPSA
    {
        static bool isLabelsPrepared = false;
        public static bool shouldInsertPicture = true;
        static string LabelCaption;
        static object oLabelCaption;
        static WordHelper wordHelper;
        static List<PictureData> picturesData;

        public static void GenerateFigureAppendix ()
        {

            if (!IsUserWantsToGenerateApp())
            {
                //Console.Clear();
                Console.WriteLine("Нажмите любую клавишу для выхода из программы...");
                return;
            }

            Console.Write("Вставлять картинки и подписи к ним? (y/n). Если выбрать 'n', то будут вставлены только подписи, без картинок: ");
            string userInput = Console.ReadLine();
            shouldInsertPicture = userInput.ToLower() == "y" ? true : false;

            ShowAbilityToChangeDictionaryMessage();
            SetLabelCaption();
            InitializeWordHelper();
            FillDocumentWithPictures();
            ReplaceUoMStyle();
            CreateTableWithReferences();
            Save();
        }

        public static void InitializeFileManager ()
        {
            Console.Clear();
            //Указываем общую папку, рабочую дирректорию, имя EXCEL-файла с набором ключей
            FileManager.GlobalPath = FileManager.SetGlobalPath();
            FileManager.ResultsFolder = FileManager.SetResultFolder();

            Console.Write("Желаете сгенерировать базу с названиями картинок? y/n: ");
            string userInput = Console.ReadLine();

            isLabelsPrepared = userInput.ToLower() == "y" ? false : true;

            if (isLabelsPrepared)
            {
                Console.Write("Укажите название базы с картинками (без .txt): ");
                userInput = Console.ReadLine();

                FileManager.CaptionsFileName = $"{userInput}.txt";
            }

            else
            {
                FileManager.DictionaryFileName = FileManager.SetDictionaryFileName();
            }
               
        }

        public static void InitializePictureData ()
        {                      
            picturesData = new List<PictureData>();

            if (isLabelsPrepared)
            {
                InputData.SetInputData(FileManager.CaptionsFileName);

                for (int index = 0; index < InputData.FileNames.Count; index++)
                {
                    PictureData picData = new PictureData();
                    picData.PictureName = InputData.FileNames[index];
                    picData.FullCaption = InputData.CaptionNames[index];
                    picturesData.Add(picData);
                }
            }
            else
            {
                //Зачитываем все данные из эксельки в класс InputData а так же определяемся с языком для вывода названия картинок
                InputData.SetInputData(FileManager.GlobalPath, FileManager.DictionaryFileName);
                InputData.language = InputData.GetLanguage();

                string combinationType = PictureData.GetCombinationType(InputData.CombinationTypes);                
                string prefix = PictureData.GetPicPrefix();

                for (int index = 0; index < InputData.FileNames.Count; index++)
                {
                    PictureData picData = new PictureData();
                    picData.PictureName = InputData.FileNames[index];
                    picData.CombinationType = combinationType;
                    picData.Prefix = prefix;

                    if (picData.PictureName.Contains('_'))
                    {
                        string[] subString = picData.PictureName.Split('_');

                        picData.SubConstruction = PictureData.GetSubConstruction(subString);
                        picData.ResultType = PictureData.GetResultType(subString);
                        picData.ConstructionType = PictureData.GetConstructionType(subString);
                        picData.ElevationTypeOf = PictureData.GetElevationType(subString);
                        picData.Elevation = PictureData.GetElevation(subString, picData);
                        picData.FullCaption = PictureData.CombineCaptionOfPic(picData);
                    }
                    else
                    {
                        picData.FullCaption = PictureData.CombineCaptionOfPic(picData, picData.PictureName);
                    }

                    picturesData.Add(picData);
                }

                Console.Clear();

                FileManager.CaptionsFileName = $"captions_{FileManager.ResultsFolder.Replace("\\", "")}.txt";
                FileManager.WriteCSV(FileManager.GlobalPath + FileManager.CaptionsFileName, picturesData);
                Console.Write(@"Файл ");
                Console.ForegroundColor = ConsoleColor.DarkYellow;
                Console.Write($"{FileManager.GlobalPath + FileManager.CaptionsFileName}");
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine(" с названиями картинок успешно создан.");
                Console.Write("Для продолжения нажмите любую клавишу...");

                Console.ReadKey();
            }
        }

        static bool IsUserWantsToGenerateApp ()
        {
            int generateWordDoc = 0;

            do
            {
                Console.Clear();
                Console.Write("Сгенерировать приложение для документа Word? (");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Write("1 - да");
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.Write("; ");
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.Write("2 - нет");
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.Write("): ");
                int.TryParse(Console.ReadLine(), out generateWordDoc);
            }
            while (generateWordDoc != 1 && generateWordDoc != 2);

            return generateWordDoc == 1 ? true : false;
        }

        static void ShowAbilityToChangeDictionaryMessage ()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"При необходимости поменяйте порядок следования картинок в Файле \"{FileManager.CaptionsFileName}\"");
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.Write("Для продолжения нажмите любую клавишу...");
            Console.ReadKey(true);
            Console.Clear();
        }

        static void SetLabelCaption ()
        {
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.Write("Введите основное название подписи к рисункам (напр. ");
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.Write("Рисунок А.");
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.Write("): ");
            LabelCaption = Console.ReadLine();
        }

        static void InitializeWordHelper ()
        {
            wordHelper = new WordHelper();

            wordHelper.OpenApplication();
            wordHelper.SetVisibilityOfApp(false);
            wordHelper.CreateDoc();
            wordHelper.AddLabelCaptionToDoc(LabelCaption, ref oLabelCaption);
        }

        static void ReplaceUoMStyle ()
        {
            wordHelper.SetVisibilityOfApp(true);
            Console.Clear();
            Console.Write("Желаете заменить все числа со степеню в формате \"^2\" на надстрочный символ? (1 - да: 2 - нет): ");

            int userInput = int.Parse(Console.ReadLine());
            if (userInput == 1) wordHelper.ReplaceUoMToSuperScript(LabelCaption.Length);
        }

        static void FillDocumentWithPictures ()
        {
            List<object> oLabelTexts = new List<object>();
            Dictionary<string, string> picCaptions = new Dictionary<string, string>();
            FileManager.ReadCSV(FileManager.GlobalPath + FileManager.CaptionsFileName, ref picCaptions);

            Console.Clear();
            foreach (var picCaption in picCaptions)
            {
                oLabelTexts.Add(" – " + picCaption.Value);
            }

            for (int index = 0; index < picCaptions.Count; index++)
            {
                string pathForPics = FileManager.GlobalPath + FileManager.ResultsFolder;
                string fileToInsert = pathForPics + picCaptions.ElementAt(index).Key;

                if (shouldInsertPicture)
                    wordHelper.InsertPicture(fileToInsert);

                wordHelper.InsertCaptionForPic(oLabelCaption, oLabelTexts[index]);
            }

            //Удаляем пробел перед номером рисунка в названии (Рисунок A. 1 -> Рисунок A.1)
            wordHelper.RemoveSpacesFromLabel(LabelCaption);
        }

        static void CreateTableWithReferences ()
        {
            Console.Clear();
            Console.Write("Желаете создать в начале документа таблицу ссылок на рисунки? (1 - да: 2 - нет): ");

            int userInput = int.Parse(Console.ReadLine());
            if (userInput == 1)
            {
                wordHelper.MoveToBeginningOfTheDoc();
                wordHelper.App.Selection.InsertParagraphBefore();
                wordHelper.MoveToBeginningOfTheDoc();
                wordHelper.InsertTable(picturesData.Count, 2, 5.0f, 12.0f);

                wordHelper.InsertCrossReferenceInTablePSA(picturesData, ref oLabelCaption);
            }
        }

        static void Save ()
        {
            Console.WriteLine("Документ готов! Для сохранения нажмите любую кнопку.");
            Console.ReadKey(true);
            Console.Write("Сохраняем документ...");

            string path = (FileManager.GlobalPath + $"Appendix_{FileManager.ResultsFolder.Replace("\\", "")}");
            string name = FileManager.ResultsFolder.Replace("\\", "");
            wordHelper.SaveDoc(path, name);
        }
    }
}
