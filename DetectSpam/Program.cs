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
        static string MailRoot = "eric@thejcrew.net";
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

                _mailApp = new Application();
                Outlook.NameSpace ns = _mailApp.GetNamespace("MAPI");
                foreach(Outlook.Account account in ns.Accounts)
                {
                    Console.WriteLine($"Found account: {account.DisplayName}");
                    MailRoot = account.DisplayName;
                    break;
                }

                foreach(var folderPath in Configuration.SpamFolderPaths)
                {
                    ScanFolder(GetFolder(folderPath));
                }
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
        /// Get a folder object based on the path
        /// </summary>
        /// -------------------------------------------------------------------------------------
        static Outlook.MAPIFolder GetFolder(string folderPath)
        {
            folderPath = MailRoot + "\\" + folderPath;
            Outlook.MAPIFolder outputFolder = null;
            foreach (var subFolder in folderPath.Split('\\'))
            {
                if (outputFolder == null) outputFolder = (Outlook.MAPIFolder)_mailApp.ActiveExplorer().Session.Folders[subFolder];
                else
                {
                    foreach(Outlook.Folder folder in outputFolder.Folders)
                    {
                        Debug.WriteLine(folder.FolderPath);
                    }
                    outputFolder = outputFolder.Folders[subFolder];
                }
            }
            return outputFolder;
        }

        /// -------------------------------------------------------------------------------------
        /// <summary>
        /// Recursively retrieves all mail messages under the starting folder
        /// </summary>
        /// -------------------------------------------------------------------------------------
        public static void ScanFolder(Outlook.MAPIFolder folder)
        {
            var subjects = new Dictionary<string, int>();
            Console.Write("Scanning Subjects ...");
            int count = 0;
            foreach (object obj in folder.Items)
            {
                if (!(obj is Outlook.MailItem)) continue;
                var item = obj as Outlook.MailItem;
                count++;
                if (count % 10 == 0) Console.Write('.');
                var subject = item.Subject;
                if (subject == null) subject = "";
                subject = subject.ToLower().Trim();
                if (!subjects.ContainsKey(subject)) subjects.Add(subject, 0);
                subjects[subject] += 1;
            }
            Console.WriteLine();

            var moveThese = new Stack<MoveData>();
            foreach (object obj in folder.Items)
            {
                if (!(obj is Outlook.MailItem)) continue;
                var item = obj as Outlook.MailItem;
                if (TryMove(item, moveThese))
                {
                    continue;
                }
                item.UnRead = false;
                var status = "na";
                status = SpamStatus(item, subjects);
                item.VotingResponse = status;


                var subject = item.Subject;
                Console.WriteLine("[" + status + "] " + subject);
                try
                {
                    item.Close(Outlook.OlInspectorClose.olSave);
                }
                catch (Exception e)
                {
                    PrintError("Error: " + e.Message);
                }
            }

            while(moveThese.Count > 0)
            {
                var moveInfo = moveThese.Pop();
                if (moveInfo.MoveTo == "")
                {
                    Console.WriteLine("***    DELETING " + moveInfo.MoveMe.Subject);
                    moveInfo.MoveMe.Delete();
                }
                else
                {
                    Console.WriteLine($"***    MOVING to {moveInfo.MoveTo}: " + moveInfo.MoveMe.Subject);
                    moveInfo.MoveMe.Move(GetFolder(moveInfo.MoveTo));
                    moveInfo.MoveMe.Close(Outlook.OlInspectorClose.olSave);
                }

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

        static List<string[]> _folderMoves;

        class MoveData
        {
            public string MoveTo;
            public Outlook.MailItem MoveMe;
        }

        /// -------------------------------------------------------------------------------------
        /// <summary>
        /// Try to move this item to another folder (or delete it)
        /// </summary>
        /// -------------------------------------------------------------------------------------
        private static bool TryMove(Outlook.MailItem item, Stack<MoveData> moveThese)
        {
            try
            {
                var searchText = $" {item.SenderEmailAddress} {item.SenderName} {item.Subject} ".ToLower();

                foreach (var moveRule in Configuration.MoveRules)
                {
                    if (searchText.Contains(moveRule.HeaderText))
                    {
                        moveThese.Push(new MoveData() { MoveTo = moveRule.TargetFolder, MoveMe = item });
                        return true;
                    }
                }
            }
            catch(Exception e)
            {
                PrintError("Error: " + e.Message);
            }

            return false;
        }



        //private static void WriteData()
        //{
        //    var config = new Configuration();
        //    config.ItemsToMove =
        //        _folderMoveList
        //        .Split('\n')
        //        .Where(i => !string.IsNullOrEmpty(i.Trim()))
        //        .Select(i =>
        //        {
        //            var parts = i.Split(',');
        //            return new MoveListItem() { HeaderText = parts[0], TargetFolder = parts[1] };
        //        })
        //        .ToArray();

        //    config.WordScores =
        //        _badWordList
        //        .Split('\n')
        //        .Where(i => !string.IsNullOrEmpty(i.Trim()))
        //        .Select(i =>
        //        {
        //            var parts = i.Split(',');
        //            return new WordScore() { RegExp = parts[0], Score = int.Parse(parts[1]) };
        //        })
        //        .ToArray();

        //    config.WhiteListTextPatterns = _goodWordList;
        //    config.WhiteListHtmlPatterns = _goodWordList;

        //    File.WriteAllText("DetectSpam.config", JsonConvert.SerializeObject(config));
        //}



        // html begins with <div
        /// -------------------------------------------------------------------------------------
        /// <summary>
        /// LooksLikeSpam
        /// </summary>
        /// -------------------------------------------------------------------------------------
        private static string SpamStatus(Outlook.MailItem item, Dictionary<string, int> subjects)
        {
            var spamStatus = "";


            var subject = item.Subject;
            if (subject == null) subject = "";
            subject = subject.Trim().ToLower();

            double score = 0;
            if(subjects.ContainsKey(subject) && subjects[subject] > 1)
            {
              //  Console.WriteLine("  Repeated subject: " + subjects[subject] );
                score += Math.Pow(5,(subjects[subject]-1));
            }
            int longRepeatedCharacterCount = 0;
            var lowerBody = subject;
            if (item.Sender != null)
            {
                var senderPart = $" {item.SenderEmailAddress} {item.SenderName} ";
                lowerBody += senderPart.ToLower();
            }
            if (item.Body != null && item.Body.Trim() != "")
            {
                lowerBody += item.Body.Trim().ToLower();
            }
            else
            {
                score += 15;
                Console.WriteLine("  No body");
            }
            var longestUnbroken = 0;
            var unbroken = 0;

            char lastChar = '\0';
            int lastCharCount = 0;
            for (int i = 0; i < lowerBody.Length; i++)
            {
                var c = lowerBody[i];
                if(c == ' ')
                {
                    //ignore spaces
                }
                else if(c == lastChar)
                {
                    lastCharCount++;
                    if(lastCharCount == 10)
                    {
                        longRepeatedCharacterCount++;
                    }
                }
                else
                {
                    lastCharCount = 0;
                    lastChar = c;
                }

                if(c == '<' || c == '\n')
                {
                    unbroken = 0;
                }
                else
                {
                    unbroken++;
                    if (unbroken > longestUnbroken) longestUnbroken = unbroken;
                }
            }
            //if (longRepeatedCharacterCount > 0) Console.WriteLine("  LongRepeats: " + longRepeatedCharacterCount);
            score += 10 * longRepeatedCharacterCount;

            //if (longestUnbroken > 500) Console.WriteLine("  LongestUnbroken: " + longRepeatedCharacterCount);
            score += (longestUnbroken / 500) * 7;

            foreach(var wordScoreRule in Configuration.WordScoreRules)
            {
                var bodyCount = Regex.Matches(lowerBody, wordScoreRule.RegExp, RegexOptions.IgnoreCase).Count;
                if(bodyCount > 0)
                {
                    score +=  bodyCount * wordScoreRule.Score;
                    //Console.WriteLine("  Suspicious: " + word + " " + bodyCount);
                }
            }

            score += GetReturnAddressScore(item);

            var htmlBody = ""; 
            if (item.HTMLBody != null) htmlBody = item.HTMLBody.Trim() ;
            var lightColorText = new Regex("[\\\" ]color[:=][\\\" ]#[ef].[ef].[ef]", RegexOptions.IgnoreCase);
            if (lightColorText.IsMatch(htmlBody))
            {
                Console.WriteLine("  Light Text");
                score += 20;
            }

            var hrefWrappedImg = new Regex("<a href.*\\\"><img", RegexOptions.IgnoreCase);
            if (hrefWrappedImg.IsMatch(htmlBody))
            {
                Console.WriteLine("  IMG With Link");
                score += 20;
            }

            if (score > 999) score = 999;
            spamStatus = score.ToString("000");
            var prefix = "OK";
            if(score >= Configuration.OKCutoffScore)
            {
                prefix = "SPAM";
            }

            var checkPrefix = "";
            foreach (var word in Configuration.WhiteListTextPatterns)
            {
                if (lowerBody.Contains(word.ToLower()))
                {
                    prefix = "GW_" + word;
                    break;
                }
            }
            foreach (var word in Configuration.WhiteListHtmlPatterns)
            {
                if (htmlBody.Contains(word))
                {
                    prefix = "GH";
                    break;
                }
            }
            if (htmlBody.StartsWith("<div ")
                || htmlBody.StartsWith("<p ")
                || htmlBody.StartsWith("<pre ")
                || !htmlBody.StartsWith("<"))
            {
                prefix = "HT";
            }


            return checkPrefix + prefix + spamStatus;

        }

        private static double GetReturnAddressScore(Outlook.MailItem item)
        {
            int symbolCount = 0;
            int CapitalCount = 0;
            int numberCount = 0;
            int score = 0;

            if (item.SenderName == null) return 30;
            if (item.SenderName.StartsWith("**")) score += 20;
            if (item.SenderName.EndsWith("**")) score += 20;

            foreach (char c in item.SenderName)
            {
                if(char.IsLetter(c))
                {
                    if (char.IsUpper(c)) CapitalCount++;
                }
                else
                {
                    if (char.IsNumber(c)) numberCount++;
                    else if (c != ' ') symbolCount++;
                }
            }

            score += symbolCount * symbolCount;
            if (CapitalCount > 2) score += CapitalCount * 2;
            if (numberCount > 2) score += numberCount * 2;
            return score;
        }

        /// -------------------------------------------------------------------------------------
        /// <summary>
        /// GetFolders
        /// </summary>
        /// -------------------------------------------------------------------------------------
        private static List<string> GetFolders(Outlook.MAPIFolder inBox)
        {
            var output = new List<string>();
            var q = new Queue<Outlook.MAPIFolder>();
            q.Enqueue(inBox);
            while(q.Count > 0)
            {
                var folder = q.Dequeue();
                output.Add(folder.FolderPath);
                Console.WriteLine(folder.FolderPath);
                foreach (Outlook.MAPIFolder subFolder in folder.Folders) q.Enqueue(subFolder); 
            }
            return output;
        }
    }
}
