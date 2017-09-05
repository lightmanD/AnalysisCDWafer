using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace AnalysisCDWafer
{
    public class FileAnalyiser
    {
        private FileStream file;
        private StreamReader streamReader;
        private string filesDirectories;

        private int rowCounter = 0; //для подсчета строк и дописывания в ексель файл
        private Dictionary<string, int> sourseDataDictionary = new Dictionary<string, int>();
        List<double> meansArray = new List<double>();

        Excel.Application excApp;
        Excel.Worksheet workSheet;
        Excel.Workbook workBook;


        public FileAnalyiser(string filesDirectories)
        {
            this.filesDirectories = filesDirectories;

        }

        public void ExcelFileCreator()
        {

            this.excApp = new Excel.Application();
            excApp.Visible = true;
            this.workBook = this.excApp.Workbooks.Add();
            this.workSheet = this.workBook.Worksheets[1];


            excApp.ActiveWorkbook.SaveAs(@"Book1.xls");

        }//нерабочий

        public void ExcelFileOpener()
        {
            string path = @"C:\Users\denis\source\repos\AnalysisCDWafer\AnalysisCDWafer\bin\Debug\result.xlsx";
            this.excApp = new Excel.Application();
            this.workBook = excApp.Workbooks.Open(path);
            this.workSheet = workBook.ActiveSheet;
            // this.workSheet.Cells[1, "A"] = "privet";
        }

        public void ExcelSaver()
        {
            DateTime dt = DateTime.Now;
            string path = @"C:\Users\denis\source\repos\AnalysisCDWafer\AnalysisCDWafer\bin\Debug\";

            path += dt.Day + "." + dt.Month + "." + dt.Year + "." + dt.Hour + "." + dt.Minute + "." + dt.Second;
            path += ".xls";

            this.workBook.SaveAs(path, Excel.XlSaveAsAccessMode.xlNoChange);

        }

        private void OpenFile()
        {
            this.file = new FileStream(filesDirectories, FileMode.Open, FileAccess.Read);
            this.streamReader = new StreamReader(file);
        }

        private void CloseFile()
        {
            streamReader.Close();
            file.Close();
        }

        public void WriteTenLine()
        {
            OpenFile();
            string temp = "";
            for (int i = 0; i < 11; i++)
            {
                temp = streamReader.ReadLine();
                this.workSheet.Cells[i + 1, 1] = temp;
                this.rowCounter++;
                Console.WriteLine(temp);
            }

            CloseFile();
        }

        //СРЕДНЕЕ
        public double Mean(List<double> inputArray)
        {
            return inputArray.Average();
        }
        // СИГМА
        public double Sigma(List<double> inputArray)
        {
            double sigma = 0;

            double mean = Mean(inputArray);

            foreach (double elem in inputArray)
            {
                sigma += Math.Pow(elem - mean, 2);
            }

            sigma = Math.Pow(sigma / (inputArray.Count - 1), 0.5);
            return sigma;
        }
        //размах
        public double Range(List<double> inputArray)
        {
            return inputArray.Max() - inputArray.Min();

        }
        // calculating  on waffer
       
        public void CollectionOfSourceData()
        {
            OpenFile();
            Console.WriteLine("Data assembling...");

            int no_of_mp = 0;
            int no_of_sequence = 0;
            int no_of_chip = 0;
            int group_number = 0;

            //ввод колличества групп
            while (true)
            {
                Console.WriteLine("Input number of point's group: ");
                var fileNumberRead = Console.ReadLine();
                Int32.TryParse(fileNumberRead, out group_number);

                if (no_of_mp % group_number == 0)
                {
                    this.sourseDataDictionary["group_number"] = group_number;
                    break;
                }

                else Console.WriteLine("The number of groups is not a multiple of the number of mp");
            }

            string line;
            //нахождение исходных данных
            while ((line = streamReader.ReadLine()) != null)
            {
                if (line.Contains("no_of_mp"))
                {
                    line = line.Replace("  ", string.Empty).Trim();
                    Char delimetr = ' ';
                    string[] substring = line.Split(delimetr);
                    no_of_mp = Convert.ToInt32(substring[1]);
                    this.sourseDataDictionary.Add("no_of_mp", no_of_mp);
                    break;
                }

                if (line.Contains("no_of_sequence"))
                {
                    line = line.Trim();
                    Char delimetr = ' ';
                    string[] substring = line.Split(delimetr);
                    no_of_sequence = Convert.ToInt32(substring[1]);
                    this.sourseDataDictionary.Add("no_of_sequence", no_of_sequence);


                }

                if (line.Contains("no_of_chip"))
                {
                    line = line.Replace("  ", string.Empty).Trim();
                    Char delimetr = ' ';
                    string[] substring = line.Split(delimetr);
                    no_of_chip = Convert.ToInt32(substring[1]);
                    this.sourseDataDictionary.Add("no_of_chip", no_of_chip);


                }

            }

            Console.WriteLine("no_of_mp = {0}\nno_of_sequence = {1}\nno_of_chip = {2}", no_of_mp, no_of_sequence, no_of_chip);

            foreach (var elem in sourseDataDictionary)
            {
                Console.WriteLine(elem);
            }

            //отсеивание всех стредних штрих
            while ((line = streamReader.ReadLine()) != null)
            {

                bool rulle_Mean = line.Contains("Mean'") && !line.Contains("Data");
                if (rulle_Mean)
                {

                    line = line.Replace(" ", string.Empty).Replace("nm", string.Empty).Replace(".", ",").Trim();

                    Char delimetr = ':';
                    string[] substring = line.Split(delimetr);
                    this.meansArray.Add(Convert.ToDouble(substring[2]));

                }

            }

            CloseFile();

        }

        public void CalculatingOnWafer()
        {
            Console.WriteLine("\n------------------------Wafer-------------------------\n");
            List<List<double>> arrays = new List<List<double>>();

            for (int i = 0; i < this.sourseDataDictionary["group_number"]; i++)
            {
                arrays.Add(new List<double>());

                for (int j = i; j < this.sourseDataDictionary["no_of_sequence"];
                    j += this.sourseDataDictionary["group_number"])
                {
                    arrays[i].Add(this.meansArray[j]);

                }
                var tempMean = Mean(arrays[i]);
                var tempSigma = Sigma(arrays[i]);
                var tempSweap = Range(arrays[i]);

                Console.Write("\nGroup #" + i);
                Console.Write("\nMean = {0}", tempMean);
                Console.Write("\nSigma = {0}", tempSigma);
                Console.Write("\nSweap = {0}\n ", tempSweap);

                //ExcelWaferWriter(i, arrays[i], tempMean, tempSigma, tempSweap);
            }
        }

        public void CalculationOnChip()
        {
            Console.WriteLine("\n------------------------Chips-------------------------\n");

            List<List<List<double>>> tempArrayChip = new List<List<List<double>>>();

            for (int i = 0; i < this.sourseDataDictionary["no_of_chip"]; i++)
            {
                Console.WriteLine("Chip #" + i);
                tempArrayChip.Add(new List<List<double>>());

                for (int k = 0; k < this.sourseDataDictionary["group_number"]; k++)
                {
                    Console.WriteLine("Group #" + k);
                    tempArrayChip[i].Add(new List<double>());

                    for (int j = k + i * this.sourseDataDictionary["no_of_mp"];
                        j < i * this.sourseDataDictionary["no_of_mp"] +
                        this.sourseDataDictionary["no_of_mp"]; j += this.sourseDataDictionary["group_number"])
                    {
                        tempArrayChip[i][k].Add(this.meansArray[j]);
                    }

                    foreach (var elem in tempArrayChip[i][k])
                        Console.Write(elem + " ");

                    var tempMean = Mean(tempArrayChip[i][k]);
                    var tempSigma = Sigma(tempArrayChip[i][k]);
                    var tempSweap = Range(tempArrayChip[i][k]);

                    Console.Write("\nMean = {0} ", tempMean);
                    Console.Write("\nSigma = {0} ", tempSigma);
                    Console.Write("\nSweap = {0} ", tempSweap);

                    //ExcelChipWriter(i, k, tempArrayChip[i][k], tempMean, tempSigma, tempSweap);

                    Console.WriteLine("\n");
                }
            }

        }

        private void ExcelChipWriter(int ChipNumber, int GroupNumber, List<double> inputList, double Mean, double Sigma, double Sweap)
        {
            rowCounter++;
            rowCounter++;
            this.workSheet.Cells[this.rowCounter, 1] = "ChipNumber";
            this.workSheet.Cells[this.rowCounter, 2] = ChipNumber + 1;
            rowCounter++;
            this.workSheet.Cells[this.rowCounter, 1] = "GroupNumber";
            this.workSheet.Cells[this.rowCounter, 2] = GroupNumber + 1;


            int colomnCounter = 2;
            this.rowCounter++;
            foreach (var elem in inputList)
            {
                this.workSheet.Cells[this.rowCounter, colomnCounter++] = elem;
            }

            rowCounter++;
            this.workSheet.Cells[this.rowCounter, 1] = "Mean";
            this.workSheet.Cells[this.rowCounter, 2] = Mean;
            rowCounter++;
            this.workSheet.Cells[this.rowCounter, 1] = "Sigma";
            this.workSheet.Cells[this.rowCounter, 2] = Sigma;
            rowCounter++;
            this.workSheet.Cells[this.rowCounter, 1] = "Sweap";
            this.workSheet.Cells[this.rowCounter, 2] = Sweap;


        }

        private void ExcelWaferWriter(int GroupNumber, List<double> inputList, double Mean, double Sigma, double Sweap)
        {

            rowCounter++;
            rowCounter++;
            this.workSheet.Cells[this.rowCounter, 1] = "GroupNumber";
            this.workSheet.Cells[this.rowCounter, 2] = GroupNumber + 1;

            int colomnCounter = 2;
            this.rowCounter++;
            foreach (var elem in inputList)
            {
                this.workSheet.Cells[this.rowCounter, colomnCounter++] = elem;
            }
            rowCounter++;
            this.workSheet.Cells[this.rowCounter, 1] = "Mean";
            this.workSheet.Cells[this.rowCounter, 2] = Mean;
            rowCounter++;
            this.workSheet.Cells[this.rowCounter, 1] = "Sigma";
            this.workSheet.Cells[this.rowCounter, 2] = Sigma;
            rowCounter++;
            this.workSheet.Cells[this.rowCounter, 1] = "Sweap";
            this.workSheet.Cells[this.rowCounter, 2] = Sweap;
        }

    }
}
