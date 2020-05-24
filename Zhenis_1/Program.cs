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
        static string[] significance = new string[] { "документ основной", "документ второстепенный", "общие документы" };
        static void Main(string[] args)
        {
            string path = @"C:\Users\Miras Talaspayev\Downloads\Telegram Desktop\BIG_DSM_MTX_004_RU_Матрица_процессов_Холдинга.xlsx"; // Miras.xlsx
            Proccess(path);
            /*
            Excel t = new Excel(path, 7);
            Console.WriteLine("Company name: {0}", t.worksheet.Name);
            Console.WriteLine("Document: {0}", t.Read_Cell_String(11, 1));
            Console.WriteLine(t.Contain(11, 15));
            Console.WriteLine("Position name: {0}", t.Read_Cell_String(2, 15));
            Console.WriteLine("Position significance: {0}", significance[t.Color(11, 15)]);
            */
            Console.ReadLine();
        }
        static List<Document> Proccess(string path)
        {
            List<Document> Output = new List<Document>();
            for (int i = 7; i <= 7; i++) // going through all companies
            {
                Excel Company = new Excel(path, i);
                for (int row = 3; row <= 232; row++)
                {
                    if (!Company.Contain(row, 1)) continue;
                    Document doc = new Document();
                    doc.name = Company.Read_Cell_String(row, 1);
                    doc.company = Company.worksheet.Name;
                    for (int column = 2; Company.Contain(2, column) == true; column++)
                    {
                        //if (row == 11 && column == 15) 
                            //Console.WriteLine("Company: {0}, Row: {1}, Column: {2}", i, row, column);
                        if (Company.Contain(row, column))
                        {
                            Position temp = new Position();
                            temp.name = Company.Read_Cell_String(2, column);
                            temp.significance = Company.Color(row, column);
                            
                            Console.WriteLine("Document: {0}", Company.Read_Cell_String(row, 1));
                            Console.WriteLine("Position name: {0}", temp.name);
                            Console.WriteLine("Position significance: {0}", significance[temp.significance]);
                            Console.WriteLine("Field: {0}", Company.Read_Cell_String(row, column));
                            
                            doc.pos.Add(temp);
                        }
                    }
                    Output.Add(doc);
                }
                Company.Close();
            }
            return Output;
        }

        static void test(string path)
        {
            List<Document> documents = Proccess(path);
            
            foreach(Document doc in documents)
            {
                Console.WriteLine("Company name: {0}", doc.company);
                Console.WriteLine("Document: {0}", doc.name);
                foreach(Position p in doc.pos)
                {
                    Console.WriteLine("Position name: {0}", p.name);
                    Console.WriteLine("Position significance: {0}", significance[p.significance]);
                }
                Console.WriteLine("=============================================================");
            }
            
        }
    }
}
// path excel file
// читаю каждую ячейку
// return 