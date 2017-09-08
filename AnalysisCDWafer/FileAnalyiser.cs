using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace AnalysisCDWafer
{
    public class FileAnalyiser
    {
        private FileStream _file;
        private StreamReader _streamReader;
        private string _filesDirectories;

        private int rowCounter = 0; //для подсчета строк и дописывания в Excel файл
        private Dictionary<string, int> sourseDataDic = new Dictionary<string, int>(); // для исходных данных
        private List<double> meansArray = new List<double>();

        string recipeName;

        Excel.Application excApp;
        Excel.Worksheet workSheet;
        Excel.Workbook workBook;


        public FileAnalyiser(string filesDirectories)
        {
            this._filesDirectories = filesDirectories;

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

            path += dt.Hour.ToString() + dt.Minute.ToString() + dt.Second.ToString() +
                "_" + dt.Day.ToString() + dt.Month.ToString() + dt.Year.ToString();
            path += "_"+recipeName.ToString()+"_";
            path += "W_"+ sourseDataDic["slot_no"].ToString();
            path += ".xls";

            this.workBook.SaveAs(path, Excel.XlSaveAsAccessMode.xlNoChange);
            this.workBook.Close();
            this.excApp.Quit();

        }

        private void OpenFile()
        {
            this._file = new FileStream(_filesDirectories, FileMode.Open, FileAccess.Read);
            this._streamReader = new StreamReader(_file);
        }

        private void CloseFile()
        {
            _streamReader.Close();
            _file.Close();
        }

        public List<string> ReadHeader()
        {
            List<string> headerList = new List<string>();
            OpenFile();
            string temp = "";
            for (int i = 0; i < 11; i++)
            {
                temp = _streamReader.ReadLine();
                headerList.Add(temp);
                Console.WriteLine(temp);
            }

            CloseFile();

            return headerList;
        }

        public double Mean(List<double> inputArray)
        {
            return inputArray.Average();
        }

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

        public double Range(List<double> inputArray)
        {
            return inputArray.Max() - inputArray.Min();

        }

        public void CollectionOfSourceData()
        {
            OpenFile();

            Console.WriteLine("Data assembling...");

            int no_of_mp = 0;
            int no_of_sequence = 0;
            int no_of_chip = 0;
            int slot_no = 0;
            int group_number = 0;

            //ввод колличества групп
            while (true)
            {
                Console.WriteLine("Input number of point's group: ");
                var fileNumberRead = Console.ReadLine();
                Int32.TryParse(fileNumberRead, out group_number);

                if (no_of_mp % group_number == 0)
                {
                    this.sourseDataDic["group_number"] = group_number;
                    break;
                }

                else Console.WriteLine("The number of groups is not a multiple of the number of mp");
            }

            string line;
            //нахождение исходных данных
            while ((line = _streamReader.ReadLine()) != null)
            {
                if (line.Contains("no_of_mp"))
                {
                    line = line.Replace("  ", string.Empty).Trim();
                    Char delimetr = ' ';
                    string[] substring = line.Split(delimetr);
                    no_of_mp = Convert.ToInt32(substring[1]);
                    this.sourseDataDic.Add("no_of_mp", no_of_mp);
                    break;
                }

                if (line.Contains("no_of_sequence"))
                {
                    line = line.Trim();
                    Char delimetr = ' ';
                    string[] substring = line.Split(delimetr);
                    no_of_sequence = Convert.ToInt32(substring[1]);
                    this.sourseDataDic.Add("no_of_sequence", no_of_sequence);

                }

                if (line.Contains("no_of_chip"))
                {
                    line = line.Replace("  ", string.Empty).Trim();
                    Char delimetr = ' ';
                    string[] substring = line.Split(delimetr);
                    no_of_chip = Convert.ToInt32(substring[1]);
                    this.sourseDataDic.Add("no_of_chip", no_of_chip);

                }

                if (line.Contains("slot_no"))
                {
                    line = line.Replace("       ", string.Empty).Trim();
                    Char delimetr = ' ';
                    string[] substring = line.Split(delimetr);
                    slot_no = Convert.ToInt32(substring[1]);
                    this.sourseDataDic.Add("slot_no", slot_no);

                }


            }

            Console.WriteLine("no_of_mp = {0}\nno_of_sequence = {1}\nno_of_chip = {2}", no_of_mp, no_of_sequence, no_of_chip);

            foreach (var elem in sourseDataDic)
            {
                Console.WriteLine(elem);
            }

            //отсеивание всех стредних штрих
            while ((line = _streamReader.ReadLine()) != null)
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

        public List<List<double>> CalculatingOnWafer()
        {
            Console.WriteLine("\n------------------------Wafer-------------------------\n");
            List<List<double>> meansOnWafer = new List<List<double>>();

            for (int i = 0; i < this.sourseDataDic["group_number"]; i++)
            {
                meansOnWafer.Add(new List<double>());

                for (int j = i; j < this.sourseDataDic["no_of_sequence"];
                    j += this.sourseDataDic["group_number"])
                {
                    meansOnWafer[i].Add(this.meansArray[j]);

                }
                var tempMean = Mean(meansOnWafer[i]);
                var tempSigma = Sigma(meansOnWafer[i]);
                var tempSweap = Range(meansOnWafer[i]);

                Console.Write("\nGroup #" + i);
                Console.Write("\nMean = {0}", tempMean);
                Console.Write("\nSigma = {0}", tempSigma);
                Console.Write("\nSweap = {0}\n ", tempSweap);

                //ExcelWaferWriter(i, arrays[i], tempMean, tempSigma, tempSweap);
            }
            return meansOnWafer;
        }

        public List<List<List<double>>> CalculationOnChip()
        {
            Console.WriteLine("\n------------------------Chips-------------------------\n");

            List<List<List<double>>> tempArrayChip = new List<List<List<double>>>();

            for (int i = 0; i < this.sourseDataDic["no_of_chip"]; i++)
            {
                Console.WriteLine("Chip #" + i);
                tempArrayChip.Add(new List<List<double>>());

                for (int k = 0; k < this.sourseDataDic["group_number"]; k++)
                {
                    Console.WriteLine("Group #" + k);
                    tempArrayChip[i].Add(new List<double>());

                    int no_of_mp = this.sourseDataDic["no_of_mp"];
                    for (int j = k + i * no_of_mp; j < i * no_of_mp + no_of_mp;
                        j += this.sourseDataDic["group_number"])
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
            return tempArrayChip;
        }

        // old method
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

        //old method
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

        //old method
        public void ExcelSaveHeader()
        {
            List<string> listHeader = ReadHeader();

            var i = 1;
            foreach (string elem in listHeader)
            {
                string fieldName = "";
                string fieldValue = "";
                List<string> listValues = new List<string>();

                Char delimetr = ' ';
                string tmp = elem.Replace("  ", string.Empty).Trim();

                //string[] substring = tmp.Split(delimetr);

                this.workSheet.Cells[i, 1] = elem;


                //fieldName = listValues[0];
                //fieldValue = listValues[1];

                //this.workSheet.Cells[i, 1] = fieldName;
                //this.workSheet.Cells[i, 2] = fieldValue;

                i++;
            }

            i++;

            foreach (var elem in this.sourseDataDic)
            {
                this.workSheet.Cells[i, 1] = elem.ToString();
                i++;
            }
        }

        public void ExcelSaveHeaderNew()
        {

            List<string> listHeader = ReadHeader();
            String pattern = @"\S+";
            List<string> matches = new List<string>();

            foreach (var expression in listHeader)
                foreach (Match m in Regex.Matches(expression, pattern))
                {
                    matches.Add(m.ToString());
                }

            matches.Remove(">ver");
            matches.Remove("MF01");
            matches.Remove("00.00");

            this.recipeName = matches[7];

            for (int i = 0; i < matches.Count; i++)
            {
                if (i % 2 == 0)
                {
                    this.workSheet.Cells[i / 2 + 1, 1] = matches[i];
                }
                else
                {
                    this.workSheet.Cells[i / 2 + 1, 2] = matches[i];

                }

            }

            int j = 12;
            foreach (var elem in this.sourseDataDic)
            {
                this.workSheet.Cells[j, 1] = elem.Key;
                this.workSheet.Cells[j, 2] = elem.Value;
                j++;
            }

        }

        public void ExcelWaferSaver(List<List<double>> inputList)
        {

            this.workSheet.Cells[19, 3] = "Group number";
            this.workSheet.Cells[20, 3] = "Mean";
            this.workSheet.Cells[21, 3] = "Sigma";
            this.workSheet.Cells[22, 3] = "Range";

            this.workSheet.Cells[24, 3] = "All values";

            var groupNum = 4;
            foreach (var group in inputList)
            {

                var mean = Mean(group);
                var sigma = Sigma(group);
                var range = Range(group);

                this.workSheet.Cells[19, groupNum] = groupNum - 3;
                this.workSheet.Cells[20, groupNum] = Mean(group);
                this.workSheet.Cells[21, groupNum] = Sigma(group);
                this.workSheet.Cells[22, groupNum] = Range(group);


                var rowCounter = 24;
                foreach (var elem in group)
                {
                    this.workSheet.Cells[rowCounter, groupNum] = elem;
                    rowCounter++;
                }
                groupNum++;
            }
        }
    }
}
