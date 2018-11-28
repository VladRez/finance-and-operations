using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HDDIRTEST
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] folder = System.IO.Directory.GetFiles(@"D:\\folder", "*.*", System.IO.SearchOption.AllDirectories);
            System.IO.StreamWriter file = new System.IO.StreamWriter("C:\\Path\\file.csv", true);
            foreach (string item in folder)
            {
                
                
                file.WriteLine(item.Replace("\\",","));
                

            }
            file.Close();
            System.Console.WriteLine("done...");
        }
    }
}
