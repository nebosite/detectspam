using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Threading;
using Application = Microsoft.Office.Interop.Outlook.Application;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.IO;
using Newtonsoft.Json;
using System.Net.NetworkInformation;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace DetectSpam
{
    class InteropStuff
    {
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        public static void HideConsole()
        {
            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_HIDE);
        }
    }

    class Program
    {
        static Application _mailApp;
        static Configuration Configuration { get; set; }
        
        //static int OKCUTOFF = 50;

        static void Main(string[] args)
        {
            var pauseForKeyboard = true;
            var continuous = false;
            var cycleMinutes = 10;
            //WriteData();  return;

            try
            {
                if (args.Length == 0) throw new ApplicationException("USAGE: detectspam (configFileName) [options]");
                var configFileName = args[0];
                for(int i = 1; i< args.Length; i++)
                {
                    var arg = args[i].ToLower().TrimStart(new char[] { '-', '/', '\\' });
                    var parts = arg.Split(new char[] { '=', ':' }, 2);
                    var name = parts[0];
                    var value = parts.Length > 1 ? parts[1] : null;
                    switch(arg) 
                    {
                        case "scripted":
                            //InteropStuff.HideConsole();
                            pauseForKeyboard = false; 
                            break;
                        case "continuous":
                            continuous = true;
                            if (value != null) cycleMinutes = int.Parse(value);
                            break;
                    }

                }
                Console.WriteLine(">>>>>>> DETECTING SPAM <<<<<<<<<<");
                if(!File.Exists(configFileName)) throw new ApplicationException("Config file does not exist: " + configFileName);

                Configuration = JsonConvert.DeserializeObject<Configuration>(File.ReadAllText(configFileName));

                var mailRoot = "eric@thejcrew.net";
                _mailApp = new Application();
                Outlook.NameSpace ns = _mailApp.GetNamespace("MAPI");
                foreach(Outlook.Account account in ns.Accounts)
                {
                    Console.WriteLine($"Found account: {account.DisplayName}");
                    mailRoot = account.DisplayName;
                    break;
                }

                while(true)
                {
                    Console.WriteLine("BEGIN SCAN ----------------");
                    var scanner = new MailScanner(Configuration, _mailApp, mailRoot);
                    foreach (var folderPath in Configuration.NotSpamFolderPaths)
                    {
                        scanner.ScanNonSpam(folderPath);
                    }

                    foreach (var folderPath in Configuration.SpamFolderPaths)
                    {
                        scanner.ScanFolder(folderPath);
                    }

                    Console.WriteLine(scanner.GetTextResults());

                    if (!continuous) break;
                    for(int i = cycleMinutes; i >= 1; i--)
                    {
                        Console.WriteLine($"Running again in {i} minutes ...");
                        Thread.Sleep(TimeSpan.FromMinutes(1));
                    }


                }
                _mailApp = null;
            }
            catch (Exception ex)
            {
                PrintError("ERROR: " + ex.ToString());
                Console.WriteLine("Note: you cannot run this as administrator.");

            }
            Console.WriteLine("DONE");
            if (pauseForKeyboard)
            {
                Console.WriteLine("Press a key to exit...");
                Console.Read();
            }
        }

        /// -------------------------------------------------------------------------------------
        /// <summary>
        /// Print some red text
        /// </summary>
        /// -------------------------------------------------------------------------------------
        static void PrintError(string message)
        {
            var originalColor = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(message);
            Console.ForegroundColor = originalColor;
        }


    }
}
