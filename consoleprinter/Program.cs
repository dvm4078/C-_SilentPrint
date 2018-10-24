using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Vls.Printer;

namespace consoleprinter
{
    class Program
    {
        static void Main(string[] args)
        {
            GetPrinterStatus inst = new GetPrinterStatus();
            foreach (KeyValuePair<string, bool> entry in (Dictionary<string, bool>)inst.GetPrinters())
            {
                Console.WriteLine("{0}=>{1}", entry.Key, entry.Value);
            }
            Console.ReadLine();
        }
    }
}
