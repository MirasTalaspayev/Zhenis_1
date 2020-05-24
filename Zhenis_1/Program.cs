using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Zhenis_1
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"C:\Users\Miras Talaspayev\Downloads\Telegram Desktop\Копия BIG_DSM_MTX_004_RU_Матрица_процессов_Холдинга.xlsx"; // Miras.xlsx
            GetDocuments(path);
        }
        static List<Document> GetDocuments(string path)
        {
            List<Document> Output = new List<Document>();
            for (int i = 1; i <= 7; i++) // going through all companies
            {
                Excel Company = new Excel(path, i);
                for (int row = 3; Company.Contain(row, 1) == true; row++) // going through all documents
                {
                    Document doc = new Document();
                    doc.name = Company.Read_Cell_String(row, 1);
                    doc.company = Company.worksheet.Name;
                    for (int column = 2; Company.Contain(2, column) == true; column++) // going through all positions
                    {
                        if (Company.Contain(row, column)) // checks for '+', is there or not
                        {
                            Position temp = new Position();
                            temp.name = Company.Read_Cell_String(2, column);
                            temp.significance = Company.Color(row, column);
                            doc.pos.Add(temp);
                        }
                    }
                    Output.Add(doc);
                }
                Company.Close();
            }
            return Output;
        }
    }
}
