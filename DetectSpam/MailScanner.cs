using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Application = Microsoft.Office.Interop.Outlook.Application;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace DetectSpam
{
    public class MailScanner
    {
        string MailRoot;
        Application _mailApp;
        Configuration Configuration { get; set; }
        Dictionary<string, int> wordPhraseCounts = new Dictionary<string, int>();

        public MailScanner(Configuration config, Application mailApp, string mailRoot)
        {
            this.MailRoot = mailRoot;
            this._mailApp = mailApp;
            this.Configuration = config;
        }

        /// -------------------------------------------------------------------------------------
        /// <summary>
        /// Get a folder object based on the path
        /// </summary>
        /// -------------------------------------------------------------------------------------
        Outlook.MAPIFolder GetFolder(string folderPath)
        {
            folderPath = MailRoot + "\\" + folderPath;
            Outlook.MAPIFolder outputFolder = null;
            foreach (var subFolder in folderPath.Split('\\'))
            {
                if (outputFolder == null) outputFolder = (Outlook.MAPIFolder)_mailApp.ActiveExplorer().Session.Folders[subFolder];
                else
                {
                    foreach (Outlook.Folder folder in outputFolder.Folders)
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
        public void ScanFolder(string folderPath)
        {
            var folder = GetFolder(folderPath);

            // Count the subjects (repeated subject look like spam)
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

            // Use Move rules to automatically move some messages to other folders
            var moveThese = new Stack<MoveData>();
            foreach (object obj in folder.Items)
            {
                if (!(obj is Outlook.MailItem)) continue;
                var item = obj as Outlook.MailItem;

                try
                {
                    if (ApplyMoveRules(item, moveThese))
                    {
                        continue;
                    }
                    item.UnRead = false;
                    var status = "na";
                    status = SpamStatus(item, subjects);
                    item.VotingResponse = status;


                    var subject = item.Subject;
                    Console.WriteLine("[" + status + "] " + subject);

                    if(!string.IsNullOrEmpty(Configuration.DefaultOutputFolder))
                    {
                        moveThese.Push(new MoveData() { MoveTo = Configuration.DefaultOutputFolder, MoveMe = item });
                    }
                }
                finally
                {
                    try
                    {
                        item.Close(Outlook.OlInspectorClose.olSave);
                    }
                    catch (Exception e)
                    {
                        PrintError("Error: " + e.Message);
                    }
                }
            }

            while(moveThese.Count > 0)
            {
                var moveInfo = moveThese.Pop();
                try
                {
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
                catch(Exception e)
                {
                    PrintError("Moving Problem: " + e.Message);
                }


            }
        }

        List<string[]> _folderMoves;

        class MoveData
        {
            public string MoveTo;
            public Outlook.MailItem MoveMe;
        }

        /// -------------------------------------------------------------------------------------
        /// <summary>
        /// Keep track of words and phrases
        /// </summary>
        /// -------------------------------------------------------------------------------------
        void analyzeText(string searchText)
        {
            var localMap = new Dictionary<string, int>();
            var list = new List<string>();
            var thisWord = new StringBuilder();
            for(int i = 0; i <= searchText.Length; i++)
            {
                if (i == searchText.Length || !char.IsLetter(searchText[i]))
                {
                    if(thisWord.Length > 0)
                    {
                        list.Add(thisWord.ToString());
                        thisWord.Clear();
                    }
                }
                else
                {
                    thisWord.Append(searchText[i]);
                }
            }
            
            void addWord(string word) {               
                if(!localMap.ContainsKey(word))
                {
                    localMap[word] = 1;

                    if(!wordPhraseCounts.ContainsKey(word))
                    {
                        wordPhraseCounts[word] = 0;
                    }
                    wordPhraseCounts[word]++;
                }
            }

            for(int i = 0; i < list.Count; i++)
            {
                addWord(list[i]);
                if (i < list.Count - 1) addWord(list[i] + " " + list[i + 1]);
                if (i < list.Count - 2) addWord(list[i] + " " + list[i + 1] + " " + list[i + 2]);
            }
        }

        /// -------------------------------------------------------------------------------------
        /// <summary>
        /// If this message matched the move rules, add it to the movable list
        /// </summary>
        /// -------------------------------------------------------------------------------------
        private bool ApplyMoveRules(Outlook.MailItem mailItem, Stack<MoveData> moveThese)
        {
            try
            {
                var recipientText = new StringBuilder();
                foreach (Outlook.Recipient recipient in mailItem.Recipients)
                {
                    recipientText.Append(recipient.Address + " ");
                }
                Outlook.PropertyAccessor oPA;
                string propName = "http://schemas.microsoft.com/mapi/proptag/0x0065001F";
                oPA = mailItem.PropertyAccessor;
                string senderMailProperty = oPA.GetProperty(propName).ToString();

                const string PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E";
                string header = oPA.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS);

                //var searchText = $" {mailItem.SenderEmailAddress} {senderMailProperty} {mailItem.SenderName} {mailItem.Subject} {recipientText} ".ToLower();
                var searchText = header.ToLower();
                analyzeText(searchText);

                Debug.WriteLine("SEARCH: " + searchText);
                foreach (var moveRule in Configuration.MoveRules)
                {
                    if (searchText.Contains(moveRule.HeaderText))
                    {
                        var startIndex = searchText.IndexOf(moveRule.HeaderText) - 10;
                        var endIndex = startIndex + moveRule.HeaderText.Length + 20;
                        if (startIndex < 0) startIndex = 0;
                        if (endIndex >= searchText.Length) endIndex = searchText.Length - 1;
                        Console.WriteLine($"    Move rule {moveRule.HeaderText} matched on ...{searchText.Substring(startIndex, endIndex - startIndex)}...");
                        moveThese.Push(new MoveData() { MoveTo = moveRule.TargetFolder, MoveMe = mailItem });
                        return true;
                    }
                }
            }
            catch (Exception e)
            {
                PrintError("Error: " + e.Message);
            }

            return false;
        }


        // html begins with <div
        /// -------------------------------------------------------------------------------------
        /// <summary>
        /// LooksLikeSpam
        /// </summary>
        /// -------------------------------------------------------------------------------------
        private string SpamStatus(Outlook.MailItem item, Dictionary<string, int> subjects)
        {
            var spamStatus = "";


            var subject = item.Subject;
            if (subject == null) subject = "";
            subject = subject.Trim().ToLower();

            double score = 0;
            if (subjects.ContainsKey(subject) && subjects[subject] > 1)
            {
                //  Console.WriteLine("  Repeated subject: " + subjects[subject] );
                score += Math.Pow(5, (subjects[subject] - 1));
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
                if (c == ' ')
                {
                    //ignore spaces
                }
                else if (c == lastChar)
                {
                    lastCharCount++;
                    if (lastCharCount == 10)
                    {
                        longRepeatedCharacterCount++;
                    }
                }
                else
                {
                    lastCharCount = 0;
                    lastChar = c;
                }

                if (c == '<' || c == '\n')
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

            foreach (var wordScoreRule in Configuration.WordScoreRules)
            {
                var bodyCount = Regex.Matches(lowerBody, wordScoreRule.RegExp, RegexOptions.IgnoreCase).Count;
                if (bodyCount > 0)
                {
                    score += bodyCount * wordScoreRule.Score;
                    //Console.WriteLine("  Suspicious: " + word + " " + bodyCount);
                }
            }

            score += GetReturnAddressScore(item);

            var htmlBody = "";
            if (item.HTMLBody != null) htmlBody = item.HTMLBody.Trim();
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
            if (score >= Configuration.OKCutoffScore)
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

        /// -------------------------------------------------------------------------------------
        /// <summary>
        /// GetReturnAddressScore
        /// </summary>
        /// -------------------------------------------------------------------------------------
        private double GetReturnAddressScore(Outlook.MailItem item)
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
                if (char.IsLetter(c))
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
        private List<string> GetFolders(Outlook.MAPIFolder inBox)
        {
            var output = new List<string>();
            var q = new Queue<Outlook.MAPIFolder>();
            q.Enqueue(inBox);
            while (q.Count > 0)
            {
                var folder = q.Dequeue();
                output.Add(folder.FolderPath);
                Console.WriteLine(folder.FolderPath);
                foreach (Outlook.MAPIFolder subFolder in folder.Folders) q.Enqueue(subFolder);
            }
            return output;
        }

        /// -------------------------------------------------------------------------------------
        /// <summary>
        /// Print some red text
        /// </summary>
        /// -------------------------------------------------------------------------------------
        void PrintError(string message)
        {
            var originalColor = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(message);
            Console.ForegroundColor = originalColor;
        }

        class KeyItem
        {
            public string key { get; set; }
            public int count { get; set; }
        }
        /// -------------------------------------------------------------------------------------
        /// <summary>
        /// Print out interesting results
        /// </summary>
        /// -------------------------------------------------------------------------------------
        public string GetTextResults()
        {
            var sortedResults = wordPhraseCounts.Keys
                .Select(k => new KeyItem() { key = k, count = wordPhraseCounts[k] })
                .OrderByDescending((a) => a.count)
                .Where(i => i.count > 3)
                .Select(i => i.key + "," + i.count);

            return String.Join("\r\n", sortedResults);
                
        }

    }
}
