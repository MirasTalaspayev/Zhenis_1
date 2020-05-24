using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;


namespace Zhenis_1
{
    class Excel
    {
        string path;
        Application excel = new Application();
        Workbook workbook;
        public Worksheet worksheet;
        public Excel(string path, int Sheet)
        {
            this.path = path;
            workbook = excel.Workbooks.Open(path);
            worksheet = workbook.Worksheets[Sheet];
        }
        public bool Contain(int row, int column)
        {
            if (worksheet.Cells[row, column].Value2 != null)
                return true;
            return false;
        }
        public string Read_Cell_String(int row, int column)
        {
            if (worksheet.Cells[row, column].Value2 != null)
                return worksheet.Cells[row, column].Value2;
            return "\0";
        }
        public string Color(int row, int column)
        {
            string col = worksheet.Cells[row, column].Interior.Color.ToString();
            
            if (col == "255") return "Оновной"; // This is Red color
            else if (col == "65535") return "второстпеннный"; // This is Yellow color
            return "Общий"; // Any other
        }
        public void Close()
        {
            workbook.Close();
        }
    }
}
