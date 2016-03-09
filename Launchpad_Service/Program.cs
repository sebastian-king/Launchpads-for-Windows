using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Diagnostics;

namespace Launchpad_Service
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Welcome to the Launchpad Service.");

            while (true)
            {
                Console.WriteLine("Going for a loop");

                try
                {
                    Process p = new Process();
                    p.StartInfo = new ProcessStartInfo(@"C:\Users\Seb\Dropbox\VS Projects\Launchpad\Launchpad\bin\x64\Release\Launchpad.exe")
                    {
                        UseShellExecute = false
                    };

                    p.Start();
                    p.WaitForExit();
                    if (p.ExitCode == Int32.MaxValue)
                    {
                        return;
                    }
                    else
                    {
                        Console.WriteLine("ExitCode: " + p.ExitCode);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception occurred: " + e.Message);
                }
                Console.WriteLine("Loop done");
            }
        }
    }
}
