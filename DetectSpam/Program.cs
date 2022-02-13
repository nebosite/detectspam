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

namespace DetectSpam
{
    class Program
    {
        static Application _mailApp;
        static Configuration Configuration { get; set; }
        
        //static int OKCUTOFF = 50;

        static void Main(string[] args)
        {
            
            //WriteData();  return;
            try
            {
                if (args.Length == 0) throw new ApplicationException("USAGE: detectspam (configFileName)");
                var configFileName = args[0];
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

                var scanner = new MailScanner(Configuration, _mailApp, mailRoot);
                foreach(var folderPath in Configuration.SpamFolderPaths)
                {
                    scanner.ScanFolder(folderPath);
                }

                Console.WriteLine(scanner.GetTextResults());
                _mailApp = null;
            }
            catch (Exception ex)
            {
                PrintError("ERROR: " + ex.ToString());
                Console.WriteLine("Note: you cannot run this as administrator.");

            }
            Console.WriteLine("DONE");
            Console.Read();
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
