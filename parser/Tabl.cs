using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace parser
{
    public class Tabl
    {
        public static HashSet<Tabl> ForSravn { get; internal set; }
        public string Id { get; set; }
        public string Name { get; set; }
        public string Info { get; set; }
        public string Sourse { get; set; }
        public string Target { get; set; }
        public string Conf { get; set; }
        public string Integ { get; set; }
        public string Avail { get; set; }

        public static IEnumerable<Tabl> EnumerateTabl(string path)
        {
            using (var workbook = new XLWorkbook(path)) 
            {
                // 
                for (int row = 3; row <= 1000000; ++row)
                {
                    if (workbook.Worksheets.Worksheet(1).Cell(row, 1).GetValue<string>() == "")
                    {
                        break;
                    }
                    else
                    {
                        var tablica = new Tabl
                        {
                            Id = workbook.Worksheets.Worksheet(1).Cell(row, 1).GetValue<string>(),
                            Name = workbook.Worksheets.Worksheet(1).Cell(row, 2).GetValue<string>(),
                            Info = workbook.Worksheets.Worksheet(1).Cell(row, 3).GetValue<string>(),
                            Sourse = workbook.Worksheets.Worksheet(1).Cell(row, 4).GetValue<string>(),
                            Target = workbook.Worksheets.Worksheet(1).Cell(row, 5).GetValue<string>(),
                            Conf = workbook.Worksheets.Worksheet(1).Cell(row, 6).GetValue<string>().Replace("0", "нет").Replace("1", "да"),
                            Integ = workbook.Worksheets.Worksheet(1).Cell(row, 7).GetValue<string>().Replace("0", "нет").Replace("1", "да"),
                            Avail = workbook.Worksheets.Worksheet(1).Cell(row, 8).GetValue<string>().Replace("0", "нет").Replace("1", "да"),
                        };
                        //ForSravn.Add(tablica);
                        yield return tablica;
                    }
                }
            }
        }
        public override string ToString()
        {
            return $"\nИДДЕНТИФИКАТОР: {Id}\nНАИМЕНОВАНИЕ: {Name}\nИНФОРМАЦИЯ ОБ УГРОЗУ: {Info}\nИСТОЧНИК УГРОЗЫ: {Sourse}\nОБЪЕКТ ВОЗДЕЙСТВИЯ: {Target}\nНАРУШЕНИЕ КОНФИДЕНЦИАЛЬНОСТИ: {Conf}\nНАРУШЕНИЕ ЦЕЛЛОСТНОСТИ: {Integ}\nНАРУШЕНИЕ ДОСТУПНОСТИ: {Avail}\n ";
        }
    }
}
