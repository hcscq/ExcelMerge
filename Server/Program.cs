using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelMerge.Common;
namespace Server
{
    class Program
    {

        
        static void Main(string[] args)
        {
            ReceiveFiles.BeginListening(Setting.IP, Setting.Port);
            Console.WriteLine("Input Q to quit.");
            Console.ReadKey();
        }       
    }
}
