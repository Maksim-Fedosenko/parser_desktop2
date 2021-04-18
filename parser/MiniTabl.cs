using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace parser
{
    public class MiniTabl
    {
        public string Id { get; set; }
        public string Name { get; set; }

        public override string ToString()
        {
            return  $"\nИДДЕНТИФИКАТОР: {Id}\nНАИМЕНОВАНИЕ: {Name}";
        }

        public static IEnumerable<MiniTabl> MiniEnumerateTabl(string path)
        {
            //int i = 0;
            using (var workbook = new XLWorkbook(path))
            {
                for (int row = 3; row <= 1000000; row++)
                {
                    if (workbook.Worksheets.Worksheet(1).Cell(row, 1).GetValue<string>() == "")
                    {
                        break;
                    }
                    else
                    {
                        var tablica = new MiniTabl
                        {


                            Id = "УБИ." + workbook.Worksheets.Worksheet(1).Cell(row, 1).GetValue<string>(),
                            Name = workbook.Worksheets.Worksheet(1).Cell(row, 2).GetValue<string>(),

                        };
                        // i++;
                        yield return tablica;
                    }
                }


            }

        }
    }
}
