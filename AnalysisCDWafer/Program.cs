using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

                var counter = 0;
                foreach (string f in filesDirectories)
                {
                    Console.WriteLine("№ {0} : {1}", counter, f);
                    counter++;
                }

                int fileNumber;
                while (true)
                {
                    Console.WriteLine("Input file number: ");
                    var fileNumberRead = Console.ReadLine();
                    Int32.TryParse(fileNumberRead, out fileNumber);
                    if (fileNumber < counter) break;
                }

                FileAnalyiser fileAnalyiser = new FileAnalyiser(filesDirectories[fileNumber]);

                fileAnalyiser.ExcelFileOpener();

                //fileAnalyiser.ReadHeader();

                fileAnalyiser.CollectionOfSourceData();
                
                fileAnalyiser.CollectionMapPoints();

                var resultWafer = fileAnalyiser.CalculatingOnWafer();

                //fileAnalyiser.CalculationOnChip();

                fileAnalyiser.ExcelSaveHeaderNew();

                fileAnalyiser.ExcelWaferSaver(resultWafer);

                fileAnalyiser.ExcelSaver();

                Console.WriteLine("\nInput command: ");
                exit = Console.ReadLine();

            }
        }
    }
}
