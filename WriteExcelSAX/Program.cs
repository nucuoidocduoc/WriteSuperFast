using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteExcelSAX
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine("start write excel :" + DateTime.Now.ToString("h:mm:ss tt") + "\n");

            Report report = new Report();

            report.WriteDataSAX();

            Console.WriteLine("end write excel :" + DateTime.Now.ToString("h:mm:ss tt") + "\n");
            Console.ReadKey();
        }
    }
}