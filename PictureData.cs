using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AppendixGenConsole
{
    public class PictureData
    {
        public string PictureName { get; set; }
        public string ResultType { get; set; }
        public string ConstructionType { get; set; }
        public string CombinationType { get; set; }
        public string SubConstruction { get; set; }
        public string Prefix { get; set; } = "";
        public string Elevation { get; set; }
        public string FullCaption { get; set; }

        public enum ElevationType
        {
            Single,
            Ranged,
            None
        }
        public ElevationType ElevationTypeOf { get; set; }

        public static string GetPicPrefix()
        {
            Console.Clear();
            Console.Write("При необходимости добавить префикс к названию, укажите его, вставив пробел после точки, например: ");
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.Write("Тоннели 91, 92UKZ. (<- тут д.б. пробел)");
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.Write("\nПри отсуствии префикса введите \'");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("N");
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.Write("\': ");

            string userInput = Console.ReadLine();

            if (userInput.ToLower() != "n")
            {
                return userInput;
            }

            return "";
        }

        public static string GetSubConstruction(string[] _subString)
        {        
            foreach(var subConstruction in InputData.SubConstructions)
            {
                string compareString = subConstruction[(int)InputData.ExcelColumnMeaning.Name].ToLower();

                if (_subString[0].ToLower().Contains(compareString))
                {
                    string subStructNum = "";

                    if (char.IsDigit(_subString[0].Last()))
                    {
                        subStructNum = _subString[0].Last().ToString();
                    }

                    if(InputData.language == InputData.ExcelColumnMeaning.LangRU)
                    {
                        return subConstruction[(int)InputData.ExcelColumnMeaning.LangRU] + subStructNum;                       
                    }
                    else
                    {
                        return subConstruction[(int)InputData.ExcelColumnMeaning.LangEN] + subStructNum;
                    }
                }
            }

            return "";
        }

        public static string GetCombinationType(List<string[]> comboTypes)
        {
            Console.Clear();
            Console.Write("Укажите комбинацию нагрузок, при необходимости, либо выберите вариант с ");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("\"None\"");
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine(" при отсуствии таковой: ");

            int numOfComboTypes = comboTypes.Count();

            for (int comboTypeNum = 0; comboTypeNum < numOfComboTypes; comboTypeNum++)
            {
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine(comboTypeNum + " - " + comboTypes[comboTypeNum][(int)InputData.ExcelColumnMeaning.Name]);
            }

            Console.WriteLine((comboTypes.Count) + " - " + "None");

            Console.ForegroundColor = ConsoleColor.Gray;
            Console.Write("Введите число от 0 до {0}: ", numOfComboTypes);

            int userInput = 0;
            if (!Int32.TryParse(Console.ReadLine(), out userInput) || userInput > numOfComboTypes || userInput < 0) return GetCombinationType(comboTypes);
            
            if (userInput == numOfComboTypes) return "";

            else
            {
                if (InputData.language == InputData.ExcelColumnMeaning.LangRU)
                    return comboTypes[userInput][(int)InputData.ExcelColumnMeaning.LangRU];
                else
                    return comboTypes[userInput][(int)InputData.ExcelColumnMeaning.LangEN];
            }
            
        }        

        public static string GetResultType(string[] _subString)
        {
            int indexOfResult = _subString.Length - 1;
            
            return GetCaptionPart(_subString, indexOfResult, InputData.ResultTypes);
        }

        public static string GetConstructionType(string[] _subString)
        {
            int indexOfConstructionType = 0;

            if (ContainsSubConstruction(_subString)) indexOfConstructionType = 1;

            return GetCaptionPart(_subString, indexOfConstructionType, InputData.ConstructionTypes);
        }

        public static string GetCaptionPart(string[] _subString, int resultIndex, List<string[]> dataForSearch)
        {         

            foreach (var constrType in dataForSearch)
            {
                string compareString = constrType[(int)InputData.ExcelColumnMeaning.Name].ToLower();

                if (_subString[resultIndex].ToLower().Contains(compareString))
                {
                    if (InputData.language == InputData.ExcelColumnMeaning.LangRU)
                    {
                        return constrType[(int)InputData.ExcelColumnMeaning.LangRU];
                    }
                    else
                    {
                        return constrType[(int)InputData.ExcelColumnMeaning.LangEN];
                    }
                }
            }

            return "";
        }

        private static bool ContainsSubConstruction(string[] _subString)
        {
            foreach (var subConstr in InputData.SubConstructions)
            {
                if (_subString[0].ToLower().Contains(subConstr[(int)InputData.ExcelColumnMeaning.Name].ToLower())) return true;
            }

            return false;
        }

        public static string GetElevation(string[] _subString, PictureData _pictureData)
        {
            string elevation = "";
            switch(_pictureData.ElevationTypeOf)
            {
                case ElevationType.Single:
                    {
                        int indexOfElevation = 1;
                        if (_pictureData.SubConstruction != "" && _pictureData.ConstructionType != "")
                            indexOfElevation = 2;

                        if (InputData.language == InputData.ExcelColumnMeaning.LangRU)
                            elevation = $"на отметке {_subString[indexOfElevation]}";

                        else
                            elevation = $"at elevation {_subString[indexOfElevation]}";

                        break;
                    }
                case ElevationType.Ranged:
                    {
                        if (_pictureData.SubConstruction != "" && _pictureData.ConstructionType != "")
                            elevation = GetRangedElevation(_subString, _pictureData, 5, 2);                           
                        else
                            elevation = GetRangedElevation(_subString, _pictureData, 4, 1);
                        break;
                    }

            }

            if (InputData.language == InputData.ExcelColumnMeaning.LangRU)
                elevation = elevation.Replace("-", "минус ");
            else
                elevation = elevation.Replace("-", "minus ");

            return elevation;
        }

        private static string GetRangedElevation(string[] _subString, PictureData _pictureData, int _lengthToCompare, int _positionOfFirstElevation)
        {
            string elevation = "";

            string constrName = _pictureData.ConstructionType.ToLower();
            string[] namesToCheck = { "wall", "beam", "column", "стен", "балк", "колон" };
            bool containsKeyWord = false;

            foreach (var name in namesToCheck)
            {
                if (constrName.Contains(name)) containsKeyWord = true;
            }

            if (containsKeyWord && _subString.Length == _lengthToCompare)
            {
                if (InputData.language == InputData.ExcelColumnMeaning.LangRU)
                    elevation = $"между отметками {_subString[_positionOfFirstElevation]} и {_subString[_positionOfFirstElevation + 1]}";
                else
                    elevation = $"from elevation {_subString[_positionOfFirstElevation]} to {_subString[_positionOfFirstElevation + 1]}";
            }

            else
            {
                if (InputData.language == InputData.ExcelColumnMeaning.LangRU)                
                    elevation = $"на отметках {_subString[_positionOfFirstElevation]}";               
                else
                    elevation = $"at elevation {_subString[_positionOfFirstElevation]}";

                int step = 1;

                while ((_positionOfFirstElevation + step < _subString.Length - 1))
                {
                    elevation += $", {_subString[_positionOfFirstElevation + step]}";
                    step++;
                }
            }

            return elevation;
        }

        public static string CombineCaptionOfPic(PictureData _picData)
        {
            string finalCaption = "";           
                        
            if (_picData.SubConstruction != "" && _picData.ConstructionType != "")
            {
                finalCaption += _picData.SubConstruction + ". ";
            }
            
            else if (_picData.SubConstruction != "")
            {
                finalCaption += _picData.SubConstruction;
            }
            
            if (_picData.ConstructionType != "")
            {
                finalCaption += _picData.ConstructionType;
            }
                
            else if (_picData.ConstructionType == "" && _picData.SubConstruction == "")
            {
                finalCaption += _picData.PictureName.Split('_')[0];
            }

            if (_picData.Elevation != "") 
                finalCaption += " " + _picData.Elevation;

            if (_picData.ResultType != "") 
                finalCaption += ". " + _picData.ResultType;

            if (_picData.CombinationType != "") 
                finalCaption += ". " + _picData.CombinationType;            

            return _picData.Prefix + finalCaption;             
        }

        public static string CombineCaptionOfPic(PictureData _picData, string _fileName)
        {
            _fileName = _fileName.Replace("кг м^2", "кг/м^2");
            _fileName = _fileName.Replace("кг м2", "кг/м^2");
            _fileName = _fileName.Replace("т м^2", "т/м^2");
            _fileName = _fileName.Replace("т м2", "т/м^2");
            return _picData.Prefix + FileManager.RemoveExtentionFromName(_fileName);
        }        

        public static ElevationType GetElevationType(string[] _subString)
        {
            bool hasSymbol = false;
            int elevCount = 0;
            foreach (string str in _subString)
            {
                bool bStr1 = str.Contains("-");
                bool bStr2 = str.Contains("+");
                if (bStr1 || bStr2)
                {
                    elevCount++;
                    hasSymbol = true;
                }
            }

            if (!hasSymbol)
            {
                return ElevationType.None;
            }

            else if (elevCount > 1)
            {
                return ElevationType.Ranged;
            }

            else
                return ElevationType.Single;

            //if (!hasSymbol)
            //    return ElevationType.None;
            //else if (ContainsSubConstruction(_subString))
            //    return CheckForElevationType(_subString.Length, 3);
            //else
            //    return CheckForElevationType(_subString.Length, 2);
        }
    }
}
