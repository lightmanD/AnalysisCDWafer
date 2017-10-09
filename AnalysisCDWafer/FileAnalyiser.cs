using System;

using System.Collections.Generic;
using System.IO;
using System.Linq;

using System.Text.RegularExpressions;

using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

//HELLO
namespace AnalysisCDWafer
{
    public class FileAnalyiser
    {
        private FileStream _file;
        private StreamReader _streamReader;
        private string _filesDirectories;

        // для исходных данных
        private Dictionary<string, int> _sourseDataDic = new Dictionary<string, int>();
        private List<double> _meansArray = new List<double>();
        private List<double> _ctrlValues = new List<double>();

        private string _recipeName;
        private string _LotID;


        Excel.Application excApp;
        Excel.Worksheet workSheet;
        Excel.Workbook workBook;

        private XDocument _xDoc;

        public FileAnalyiser(string filesDirectories)
        {
            this._filesDirectories = filesDirectories;

        }

        //нерабочий
        public void ExcelFileCreator()
        {

            this.excApp = new Excel.Application();
            excApp.Visible = true;
            this.workBook = this.excApp.Workbooks.Add();
            this.workSheet = this.workBook.Worksheets[1] as Excel.Worksheet;


            excApp.ActiveWorkbook.SaveAs(@"Book1.xls");

        }

        public void ExcelFileOpener()
        {
            string fileResultName = @"\result.xlsx";
            string fileDirectory = Directory.GetCurrentDirectory();
            string path = fileDirectory + fileResultName;


            this.excApp = new Excel.Application();
            this.workBook = excApp.Workbooks.Open(path);
            this.workSheet = workBook.ActiveSheet as Excel.Worksheet;

        }

        public void ExcelNewSheet()
        {
            Excel.Worksheet newWorksheet;
            newWorksheet = (Excel.Worksheet)this.excApp.Worksheets.Add();
            this.workSheet = newWorksheet;
            this.workSheet = workBook.ActiveSheet as Excel.Worksheet;
        }

        public void ExcelSaver()
        {
            DateTime dt = DateTime.Now;
            string path = Directory.GetCurrentDirectory() + @"\results\";

            path += "_" + _LotID;
            path += "_" + _recipeName.ToString() + "_";
            path += "W_" + _sourseDataDic["slot_no"].ToString();
            path += ".xls";

            this.workBook.SaveAs(path, Excel.XlSaveAsAccessMode.xlNoChange);
            this.workBook.Close();
            this.excApp.Quit();

            this.workBook = null;
            this.excApp = null;

        }

        private void OpenFile()
        {
            this._file = new FileStream(_filesDirectories, FileMode.Open, FileAccess.Read);
            this._streamReader = new StreamReader(_file);
        }

        private void CloseFile()
        {
            _streamReader.Close();
            _streamReader = null;
            _file.Close();
            _file = null;
        }

        private void LoadXmlConfig()
        {
            _xDoc = new XDocument();
            _xDoc = XDocument.Load("RecipeData.xml");
        }

        private void SaveXmlConfig()
        {
            _xDoc.Save("RecipeData.xml");
            _xDoc = null;
        }

        public List<string> ReadHeadNew()
        {
            Console.WriteLine("+Считывание хэдера+");
            List<string> headerList = new List<string>();
            OpenFile();
            string temp = "";
            for (int i = 0; i < 11; i++)
            {
                temp = _streamReader.ReadLine();
                headerList.Add(temp);
                Console.WriteLine(temp);
            }

            String pattern = @"\S+";

            //String pattern = "(\".*?\")";
            List<string> matches = new List<string>();

            foreach (var expression in headerList)
                foreach (Match m in Regex.Matches(expression, pattern))
                {
                    matches.Add(m.ToString());
                }

            matches.Remove(">ver");
            matches.Remove("MF01");
            matches.Remove("00.00");

            _recipeName = matches[7].Trim();
            _LotID = matches[11].Trim().Replace("\"", string.Empty);


            CloseFile();

            headerList = null;
            return matches;
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

        public double Max(List<double> inputArray)
        {
            return inputArray.Max();
        }

        public double Min(List<double> inputArray)
        {
            return inputArray.Min();
        }

        public void CollectionOfSourceData()
        {
            Console.WriteLine("+Сбор исходнных данных+");
            OpenFile();

            string line;
            string typeOfData = "";
            //нахождение исходных данных
            while ((line = _streamReader.ReadLine()) != null)
            {
                if (line.Contains("no_of_mp"))
                {
                    line = line.Replace("  ", string.Empty).Trim();
                    Char delimetr = ' ';
                    string[] substring = line.Split(delimetr);

                    this._sourseDataDic.Add("no_of_mp", Convert.ToInt32(substring[1]));

                }

                if (line.Contains("no_of_sequence"))
                {
                    line = line.Trim();
                    Char delimetr = ' ';
                    string[] substring = line.Split(delimetr);
                    this._sourseDataDic.Add("no_of_sequence", Convert.ToInt32(substring[1]));

                }

                if (line.Contains("no_of_chip"))
                {
                    line = line.Replace("  ", string.Empty).Trim();
                    Char delimetr = ' ';
                    string[] substring = line.Split(delimetr);
                    this._sourseDataDic.Add("no_of_chip", Convert.ToInt32(substring[1]));

                }

                if (line.Contains("slot_no"))
                {
                    line = line.Replace("       ", string.Empty).Trim();
                    Char delimetr = ' ';
                    string[] substring = line.Split(delimetr);
                    this._sourseDataDic.Add("slot_no", Convert.ToInt32(substring[1]));

                }

                if (line.Contains("1 : Data"))
                {
                    line = line.Replace(" ", string.Empty).Replace("nm", string.Empty).Replace(".", ",").Trim();
                    Char delimetr = ':';
                    string[] substring = line.Split(delimetr);

                    typeOfData = substring[2];

                    break;
                }
            }


            if (typeOfData == "Mean'")
                //сбор всех стредних штрих
                while ((line = _streamReader.ReadLine()) != null)
                {

                    bool rulle_Mean = line.Contains("Mean'") && !line.Contains("Data");
                    if (rulle_Mean)
                    {

                        line = line.Replace(" ", string.Empty).Replace("nm", string.Empty).Replace(".", ",").Trim();

                        Char delimetr = ':';
                        string[] substring = line.Split(delimetr);
                        this._meansArray.Add(Convert.ToDouble(substring[2]));
                        Console.WriteLine(Convert.ToDouble(substring[2]));
                    }

                }
            //сбор всех диаметров
            else
                while ((line = _streamReader.ReadLine()) != null)
                {

                    bool rulle_Mean = line.Contains("Diameter") && !line.Contains("Data") && !line.Contains("X") && !line.Contains("Y") && !line.Contains("Object") && !line.Contains("Measurement");
                    if (rulle_Mean)
                    {

                        line = line.Replace(" ", string.Empty).Replace("nm", string.Empty).Replace(".", ",").Trim();

                        Char delimetr = ':';
                        string[] substring = line.Split(delimetr);
                        this._meansArray.Add(Convert.ToDouble(substring[2]));

                    }

                }

            CloseFile();

        }

        public void CollectionGroupNumber()
        {
            Console.WriteLine("+Сбор колличества групп+");
            int group_number = 0;
            int INT = 0;
            while (true)
            {
                Console.WriteLine("Input number of point's group: ");
                var fileNumberRead = Console.ReadLine();
                Int32.TryParse(fileNumberRead, out group_number);
                this._sourseDataDic["group_number"] = group_number;
                if (group_number.GetType().ToString() == INT.GetType().ToString())

                    break;

            }
        }

        public List<List<double>> CalculatingOnWafer()
        {
            Console.WriteLine("\n----------------Расчет-по-пластине--------------------\n");

            List<List<double>> meansOnWafer = new List<List<double>>();

            for (int i = 0; i < _sourseDataDic["group_number"]; i++)
            {
                meansOnWafer.Add(new List<double>());

                for (int j = i; j < this._sourseDataDic["no_of_sequence"];
                    j += this._sourseDataDic["group_number"])
                {
                    meansOnWafer[i].Add(this._meansArray[j]);

                }
                var tempMean = Mean(meansOnWafer[i]);
                var tempSigma = Sigma(meansOnWafer[i]);
                var tempSweap = Range(meansOnWafer[i]);

                Console.Write("\nGroup #" + i);
                Console.Write("\nMean = {0}", tempMean);
                Console.Write("\nSigma = {0}", tempSigma);
                Console.Write("\nSweap = {0}\n", tempSweap);


            }
            return meansOnWafer;
        }

        public List<List<List<double>>> CalculationOnChip()
        {
            Console.WriteLine("\n------------------------Chips-------------------------\n");

            List<List<List<double>>> tempArrayChip = new List<List<List<double>>>();

            for (int i = 0; i < this._sourseDataDic["no_of_chip"]; i++)
            {
                Console.WriteLine("Chip #" + i);
                tempArrayChip.Add(new List<List<double>>());

                for (int k = 0; k < this._sourseDataDic["group_number"]; k++)
                {
                    Console.WriteLine("Group #" + k);
                    tempArrayChip[i].Add(new List<double>());

                    int no_of_mp = this._sourseDataDic["no_of_mp"];
                    for (int j = k + i * no_of_mp; j < i * no_of_mp + no_of_mp;
                        j += this._sourseDataDic["group_number"])
                    {
                        tempArrayChip[i][k].Add(this._meansArray[j]);
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

        public void ExcelSaverHead(List<string> matches)
        {
            Console.WriteLine("+Сохранение заголовка+");
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
            foreach (var elem in this._sourseDataDic)
            {
                this.workSheet.Cells[j, 1] = elem.Key;
                this.workSheet.Cells[j, 2] = elem.Value;
                j++;
            }
        }

        public void ExcelWaferSaver(List<List<double>> inputList)
        {
            Console.WriteLine("+Запись расчетов по пластине+");



            this.workSheet.Cells[19, 3] = "Group name";
            this.workSheet.Cells[20, 3] = "Mean";
            this.workSheet.Cells[21, 3] = "Sigma";
            this.workSheet.Cells[22, 3] = "Range";
            this.workSheet.Cells[23, 3] = "Min";
            this.workSheet.Cells[24, 3] = "Max";


            this.workSheet.Cells[26, 3] = "All values";

            var groupNum = 4;
            foreach (var group in inputList)
            {

                var mean = Mean(group);
                var sigma = Sigma(group);
                var range = Range(group);



                this.workSheet.Cells[20, groupNum] = Mean(group);
                this.workSheet.Cells[21, groupNum] = Sigma(group);
                this.workSheet.Cells[22, groupNum] = Range(group);

                this.workSheet.Cells[23, groupNum] = Min(group);
                this.workSheet.Cells[24, groupNum] = Max(group);


                var rowCounter = 26;
                foreach (var elem in group)
                {
                    this.workSheet.Cells[rowCounter, groupNum] = elem;
                    rowCounter++;
                }
                groupNum++;
            }

            //писать в Excel нужно из Xml метод CollectGroupsNameFromXmlDataRecipe
            //так оно и пишет
            var mpList = CollectGroupsNameFromXmlDataRecipe();

            int columnCounter = 4;
            foreach (var elem in mpList)
            {
                this.workSheet.Cells[19, columnCounter++] = elem;
            }

        }

        public List<string> CollectionAllMapPoints()
        {
            List<string> mpList = new List<string>();
            string line;

            OpenFile();

            while ((line = _streamReader.ReadLine()) != null)
            {
                string patternMp = "~mp_name";

                if (line.Contains(patternMp))
                {
                    mpList.Add(line);
                }
            }

            CloseFile();

            List<string> mpNamesList = new List<string>();
            string re1 = ".*?"; // Non-greedy match on filler
            string re2 = "(\".*?\")";   // Double Quote String 1

            Regex r = new Regex(re1 + re2, RegexOptions.IgnoreCase | RegexOptions.Singleline);
            foreach (var expression in mpList)
            {
                Match m = r.Match(expression);
                if (m.Success)
                {
                    String string1 = m.Groups[1].ToString();
                    string1 = string1.Replace('"', ' ');
                    mpNamesList.Add(string1.ToString());
                }
            }

            return mpNamesList;
        }

        public List<string> FilteringMPNames(List<string> inputList)
        {

            foreach (var elem in inputList) Console.WriteLine(elem);

            //ввод колличества групп
            CollectionGroupNumber();

            List<string> groupNames = new List<string>();

            for (int i = 0; i < this._sourseDataDic["group_number"]; i++)
            {
                string temp = inputList[i].ToString();
                char delim = '-';
                int indexDelim = temp.IndexOf(delim);
                temp = temp.Substring(indexDelim + 1);
                groupNames.Add(temp);
            }
            return groupNames;
        }

        public bool CheckRecipeInConfig()
        {
            Console.WriteLine("+Проверка наличия рецепта+");
            LoadXmlConfig();

            foreach (XElement recipeElem in _xDoc.Element("recipes").Elements("recipe"))
            {
                XAttribute nameAttribute = recipeElem.Attribute("name");
                if (nameAttribute.Value.ToString() == _recipeName)
                {

                    Console.WriteLine("+Пройдена+");
                    return true;
                }

            }
            SaveXmlConfig();

            Console.WriteLine("+Непройдена+");
            return false;
        }

        public void FormRecipeDataFilling()
        {
            Console.WriteLine("Запись новых данных для рецепта+");
            LoadXmlConfig();

            var mpList = FilteringMPNames(CollectionAllMapPoints());
            var ctrlValues = CollectCtrlValue();

            foreach (var elem in mpList)
            {
                Console.WriteLine(elem);
            }

            XElement root = _xDoc.Element("recipes");

            root.Add(new XElement("recipe",
                new XAttribute("name", _recipeName),
                new XElement("group_number", _sourseDataDic["group_number"]),
                new XElement("groups", mpList),
                new XElement("ctrl_value", ctrlValues)));

            SaveXmlConfig();
        }

        public void CollectionDataFromXmlDataRecipe()
        {
            Console.WriteLine("+Сбор данных из RecipiData.xml+");

            LoadXmlConfig();

            foreach (XElement recipeElem in _xDoc.Element("recipes").Elements("recipe"))
            {
                XAttribute nameAttribute = recipeElem.Attribute("name");
                XElement groupsNumElement = recipeElem.Element("group_number");
                XElement groupsElement = recipeElem.Element("groups");
                XElement ctrlValuesElement = recipeElem.Element("ctrl_value");

                if (nameAttribute.Value.ToString() == _recipeName)
                {
                    Int32.TryParse(groupsNumElement.Value.ToString(), out int group_number);
                    _sourseDataDic["group_number"] = group_number;

                    _ctrlValues = CollecttionCtrlValuesFromXml(ctrlValuesElement.Value.ToString());
                }
            }
            SaveXmlConfig();
        }

        public List<string> CollectGroupsNameFromXmlDataRecipe()
        {
            Console.WriteLine("+Сбор имен групп из RecipiData.xml+");
            LoadXmlConfig();

            List<string> listGroups = new List<string>();

            foreach (XElement recipeElem in _xDoc.Element("recipes").Elements("recipe"))
            {
                XAttribute nameAttribute = recipeElem.Attribute("name");
                XElement groupsNumElement = recipeElem.Element("group_number");
                XElement groupsElement = recipeElem.Element("groups");

                if (nameAttribute.Value.ToString() == _recipeName)
                {
                    var groups = groupsElement.Value;
                    listGroups = groups.Split(' ').ToList<string>();

                }
            }

            SaveXmlConfig();
            return listGroups;
        }

        private string CollectCtrlValue()
        {
            Console.WriteLine("Введите значения границ(формат USL UCL Target LCL LSL) ");
            string input = Console.ReadLine();

            return input;
        }

        private List<double> CollecttionCtrlValuesFromXml(string inputString)
        {
            List<double> ctrlValues = new List<double>();

            string[] stringSplit = inputString.Split(' ');

            foreach (var elem in stringSplit)
            {
                ctrlValues.Add(Double.Parse(elem));
            }

            return ctrlValues;
        }

        private void ColorValue()
        {

        }
    }
}

