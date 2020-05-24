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
        public int Color(int row, int column)
        {
            string col = worksheet.Cells[row, column].Interior.Color.ToString();
            
            string color = "";
            if (col == "255") color = "Red";
            else if (col == "65535") color = "Yellow";

            if (color == ConsoleColor.Red.ToString())
                return 0;
            else if (color == ConsoleColor.Yellow.ToString())
                return 1;
            return 2;
        }
        public void Close()
        {
            workbook.Close();
        }
    }
}
