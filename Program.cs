using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AppendixGenConsole
{
    class Program
    {
        static void Main (string[] args)
        {
            Menu();
            Console.ReadKey();
        }

        static void CreatePSA ()
        {
            AppendixPSA.InitializeFileManager();
            AppendixPSA.InitializePictureData();
            AppendixPSA.GenerateFigureAppendix();
        }

        static void CreateFRS ()
        {
            AppendixFRS.MenuFRS();
        }

        static void Menu ()
        {
            Console.WriteLine("Выберите необходимое действие: ");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("1. Генерация приложений для отчета прочности");
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("2. Генерация приложения для отчета спектров отклика");
            Console.ForegroundColor = ConsoleColor.Gray;

            Console.Write("Для этого введите число ");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write("1");
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.Write(" или ");
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.Write("2");
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.Write(" и нажмите Enter: ");

            int action = 0;
            Int32.TryParse(Console.ReadLine(), out action);

            switch (action)
            {
                case 1:
                    {
                        CreatePSA();
                        break;
                    }
                case 2:
                    {
                        CreateFRS();
                        break;
                    }
            }
        }


    }
}
