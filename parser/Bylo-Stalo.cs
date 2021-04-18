using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace parser
{
    public class Bylo_Stalo
    {
        public string Было { get; set; }
        //HashSet<string> Bylo = new HashSet<string>();
        public string Стало { get; set; }
       // private HashSet<string> Stalo = new HashSet<string>();
       //private string Itog { get; set; }

        public override string ToString()
        {
            return $"\n[БЫЛО]{Было}\n[Стало]{Стало}" ;
        }

        public Bylo_Stalo(string a, string b)
        {
            Было =a;
            Стало =b;
        }
    }
}
