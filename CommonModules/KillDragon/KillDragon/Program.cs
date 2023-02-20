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
        private static readonly string KB_PROCESS_1 = "KBPro";
        private static readonly string KB_PROCESS_2 = "ProcHandler";

        static void Main(string[] args)
        {
            try
            {
                List<string> processes = new List<string>();
                processes.Add("dgnuiasvr");
                processes.Add("dgnuiasvr_x64");
                processes.Add("dragonbar");
                processes.Add("dgnria_nmhost");
                processes.Add("natspeak");
                processes.Add(KB_PROCESS_1);
                processes.Add(KB_PROCESS_2);

                Console.WriteLine("*********************************************************************************************");
                Console.WriteLine("This program will kill all processes associated with Dragon Naturallyspeaking.");
                Console.WriteLine("This is useful when Dragon freezes up / stops responding and needs to be stopped completely.");
                Console.WriteLine("Don't do this if you can close Dragon cleanly and properly, thereby saving and closing files.");
                Console.WriteLine("*********************************************************************************************");
                Console.WriteLine();

                string killsString = null;
                ConsoleKeyInfo key = new ConsoleKeyInfo();
                

                do
                {
                    killsString = GetRunningProcessListString(processes);
                    
                    if (killsString.Length == 0)
                    {
                        Console.WriteLine("There are no Dragon process running.  Press any key to close.");
                        Console.ReadKey(true);
                        return;
                    }

                    if (killsString.Contains(KB_PROCESS_1))
                    //if (killsString.Contains(KB_PROCESS_1) || killsString.Contains(KB_PROCESS_2)) // Uncomment for testing only
                    {
                        Console.WriteLine();
                        Console.WriteLine("It appears that Knowbrainer is running.  It is preferable to shutdown Knowbrainer manually if possible.");
                        Console.WriteLine("Please try to manually shutdown Knowbrainer and then press R to recheck.");
                        Console.WriteLine("Press K to skip this check and forcibly kill Knowbrainer along with Dragon.");
                        Console.WriteLine("Press X to exit.");
                        while (true)
                        {
                            key = Console.ReadKey(true);
                            if (key.Key == ConsoleKey.K)
                            {
                                break;
                            }
                            else if (key.Key == ConsoleKey.R)
                            {
                                break;
                            }
                            else if (key.Key == ConsoleKey.X)
                            {
                                return;
                            }
                            else
                            {
                                Console.WriteLine("Unrecognized entry: " + key.Key.ToString());
                            }
                        }
                    }
                    else
                    {
                        // Know brainer is not running anymore, continue on
                        key = new ConsoleKeyInfo('K', ConsoleKey.K, false, false, false);
                    }

                } while (key.Key != ConsoleKey.K);

                bool exitApp = false;
                do
                {
                    killsString = GetRunningProcessListString(processes);

                    Console.WriteLine();
                    Console.WriteLine("Here are the currently running Dragon/KB processes:");
                    Console.WriteLine();
                    Console.Write(killsString);
                    Console.WriteLine();
                    Console.WriteLine("Do you want to kill these processes?");
                    Console.WriteLine("Press K to kill the processes.");
                    Console.WriteLine("Press R to refresh the list.");
                    Console.WriteLine("Press X to exit without killing any processes.");
                    key = Console.ReadKey(true);
                    if (key.Key == ConsoleKey.X)
                    {
                        Console.WriteLine();
                        Console.WriteLine("No action taken.");
                        Console.WriteLine("Press any key to close.");
                        Console.ReadKey(true);
                        exitApp = true;
                    }
                    else if (key.Key == ConsoleKey.R)
                    {
                        continue;
                    }
                    else if (key.Key == ConsoleKey.K)
                    {
                        Console.WriteLine();
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
                        exitApp = true;
                    }
                } while (!exitApp);
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine(ex.Message);
                Console.WriteLine();
                Console.WriteLine("Press any key to close.");
                Console.ReadKey();
            }
        }

        private static string GetRunningProcessListString(List<string> checkProcesses)
        {
            StringBuilder toKillList = new StringBuilder();
            foreach (string name in checkProcesses)
            {
                Process[] processList = Process.GetProcessesByName(name);
                if (processList.Length > 0)
                {
                    toKillList.AppendLine(name);
                }
            }

            return toKillList.ToString();
        }
    }
}
