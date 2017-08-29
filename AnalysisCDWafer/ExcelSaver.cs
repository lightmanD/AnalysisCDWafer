using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace AnalysisCDWafer
{
    class ExcelSaver
    {
        private Application ObjExcel;
        ExcelSaver()
        {
            ExcelFileCreator();
        }

        void ExcelFileCreator()
        {
            ObjExcel = new Application();
            Workbook ObjWorkBook;
            Worksheet ObjWorkSheet;
        }

    }
}
