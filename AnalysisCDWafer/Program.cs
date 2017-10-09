using System;
using System.Collections.Generic;
using System.IO;


namespace AnalysisCDWafer
{
    class Program
    {
        static void Main(string[] args)
        {
            string exit = "";

            Console.ForegroundColor = ConsoleColor.Green;

            while (!(exit.Equals("exit") || exit.Equals("q")))
            {
                string[] filesDirectories = Directory.GetFiles("files/");

                int counter = 0;
                foreach (string f in filesDirectories)
                {
                    Console.WriteLine("№ {0} : {1}", counter, f);
                    counter++;
                }

                //int fileNumber;
                List<int> fileNumbers = new List<int>();
                while (true)
                {
                    Console.WriteLine("Введите номера файлов: ");
                    string fileNumberRead = Console.ReadLine().Trim();
                    string[] stringSplit = fileNumberRead.Split(' ');

                    //parse file's numbers
                    foreach (var elem in stringSplit)
                    {
                        int num = Int16.Parse(elem);
                        if (num > counter || num < 0) break;
                        fileNumbers.Add(num);

                    }

                    if (counter < fileNumbers.Count || fileNumbers.Count != 0 || stringSplit.Length == fileNumbers.Count) break;
                }

                int iter = 0;
                FileAnalyiser fileAnalyiser;

                foreach (var num in fileNumbers)
                {
                    string fileDirectory = filesDirectories[num];

                    fileAnalyiser = new FileAnalyiser(fileDirectory);

                    var headMatches = fileAnalyiser.ReadHeadNew();

                    fileAnalyiser.CollectionOfSourceData();

                    if (!fileAnalyiser.CheckRecipeInConfig())
                    {
                        fileAnalyiser.FormRecipeDataFilling();
                    }

                    fileAnalyiser.CollectionDataFromXmlDataRecipe();

                    var resultWafer = fileAnalyiser.CalculatingOnWafer();

                    //fileAnalyiser.CalculationOnChip();

                    fileAnalyiser.ExcelFileOpener();

                    fileAnalyiser.ExcelSaverHead(headMatches);

                    fileAnalyiser.ExcelWaferSaver(resultWafer);

                    fileAnalyiser.ExcelSaver();

                    iter++;
                }
                Console.WriteLine("\nInput command (q or exit to quit): ");
                exit = Console.ReadLine();
            }
        }


    }
}
