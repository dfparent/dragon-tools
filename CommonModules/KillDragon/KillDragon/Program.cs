using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KillDragon
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> processes = new List<string>();
            processes.Add("dgnuiasvr");
            processes.Add("dgnuiasvr_x64");
            processes.Add("dragonbar");
            processes.Add("dgnria_nmhost");
            processes.Add("natspeak");
            processes.Add("KBPro");
            processes.Add("ProcHandler");

            Console.WriteLine("*********************************************************************************************");
            Console.WriteLine("This program will kill all processes associated with Dragon Naturallyspeaking.");
            Console.WriteLine("This is useful when Dragon freezes up / stops responding and needs to be stopped completely.");
            Console.WriteLine("Don't do this if you can close Dragon cleanly and properly, thereby saving and closing files.");
            Console.WriteLine("*********************************************************************************************");
            Console.WriteLine();

            StringBuilder toKillList = new StringBuilder();
            foreach (string name in processes)
            {
                Process[] processList = Process.GetProcessesByName(name);
                if (processList.Length > 0)
                {
                    toKillList.AppendLine(name);
                }
            }

            if (toKillList.Length == 0)
            {
                Console.WriteLine("There are no Dragon process running.  Press any key to close.");
                Console.ReadKey(true);
                return;
            }

            Console.Write(toKillList.ToString());
            Console.WriteLine();
            Console.WriteLine("Do you want to kill these processes? (Y/N)");
            ConsoleKeyInfo key = Console.ReadKey(true);
            if (key.Key != ConsoleKey.Y)
            {
                Console.WriteLine();
                Console.WriteLine("No action taken.");
                Console.WriteLine("Press any key to close.");
                Console.ReadKey(true);
                return;
            }

            foreach (string name in processes)
            {
                foreach (Process aProcess in Process.GetProcessesByName(name))
                {
                    Console.WriteLine("Killing " + name);
                    aProcess.Kill();
                }
            }

            Console.WriteLine();
            Console.WriteLine("Done.  You may now restart Dragon.  Press any key to close.");
            Console.ReadKey();

        }
    }
}
