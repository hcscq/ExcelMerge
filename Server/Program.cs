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
            ConsoleKeyInfo consoleKeyInfo;
            while (!(consoleKeyInfo=Console.ReadKey()).Key.Equals(ConsoleKey.Escape))
            {
                switch (consoleKeyInfo.Key)
                {
                    case ConsoleKey.C:
                        Console.Clear();
                        break;
                    default:break;
                }
            }
            ReceiveFiles.Working = false;
        }       
    }
}
