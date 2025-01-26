using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Threading;
using Application = Microsoft.Office.Interop.Outlook.Application;
using Exception = System.Exception;
using System.Text.RegularExpressions;
using System.IO;
using Newtonsoft.Json;
using System.Runtime.InteropServices;
using System.Diagnostics;

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

        static void Main(string[] args) {

            if(args.Length > 0 && args[0] == "-farmcontacts") {
                FarmContacts();
            } else if (args.Length > 0 && args[0] == "-cleanjunk")
            {
                CleanJunk(args[1], args.Length > 2 && args[2].ToLower() == "delete" ? true : false);
            } else
            {
                RunSpamWatch(args);
            }
        }

        // --------------------------------------------------------------------------------
        // --------------------------------------------------------------------------------
        static void RunInMailContext(Action<Application, string, Outlook.NameSpace> runMe) 
        {
            var mailRoot = "(unknown)";
            _mailApp = new Application();
            Outlook.NameSpace mapiNameSpace = _mailApp.GetNamespace("MAPI");
            foreach(Outlook.Account account in mapiNameSpace.Accounts)
            {
                Console.WriteLine($"Found account: {account.DisplayName}");
                mailRoot = account.DisplayName;
                break;
            }

            runMe(_mailApp, mailRoot, mapiNameSpace);
            _mailApp = null;
        }

        // --------------------------------------------------------------------------------
        // class for working with groups of emails from a single sender
        // --------------------------------------------------------------------------------
        public class SenderGroup
        {
            public string Domain { get; set; }
            public string Recipient { get; set; }
            public Dictionary<string, int> Addresses = new Dictionary<string, int>();
            public string UnsubscribeLink = null;
            public DateTime UnsubscribeDate;
            public DateTime MostRecent = DateTime.Parse("1970/1/1");

            public int TotalCount
            {
                get
                {
                    var total = 0;
                    foreach (var count in Addresses.Values) total += count;
                    return total;
                }
            }

            public SenderGroup(string domain)
            {
                this.Domain = domain;
            }

            public void AddMailItem(Outlook.MailItem item)
            {
                if(item.ReceivedTime > MostRecent)
                {
                    MostRecent = item.ReceivedTime;
                }

                if(this.Recipient == null)
                { 
                    this.Recipient = item.To;
                }

                if(!this.Addresses.ContainsKey(item.SenderName))
                {
                    this.Addresses[item.SenderName] = 1;
                } else
                {
                    this.Addresses[item.SenderName]++;
                }

            }

        }

        // --------------------------------------------------------------------------------
        // --------------------------------------------------------------------------------
        static void CleanJunk(string folder, bool deleteOlderMail)
        {
            var groups = new Dictionary<string, SenderGroup>();

            string getDomainName(string emailAddress)
            {
                var parts = emailAddress.Split(".@".ToCharArray());
                if(parts.Length == 1) return parts[0];
                else return (parts[parts.Length-2] + "." + parts[parts.Length-1]).ToLower();
            }

            // --------------------------------------------------------------------------------
            // --------------------------------------------------------------------------------
            RunInMailContext((outlookApp, mailRoot, outlookNamespace) =>
            {
                var exceptionCount = 0;
                try
                {
                    var count = 0;
                    var scanner = new MailScanner(null, outlookApp, mailRoot);
                    scanner.ProcessFolder(folder, (item) =>
                    {
                        count++; 
                        if(count % 10 == 0)
                        {
                            Console.Error.Write($"{count}       \r");
                        }

                        try
                        {
                            var domain = getDomainName(item.SenderEmailAddress);

                            if (!groups.ContainsKey(domain))
                            {
                                groups.Add(domain, new SenderGroup(domain));
                            }

                            var group = groups[domain];

                            if (group.UnsubscribeLink == null || group.UnsubscribeDate < item.ReceivedTime)
                            {
                                var anchorMatches = Regex.Matches(item.HTMLBody, "(<a.*?>(.*?)</a>)", RegexOptions.IgnoreCase);
                                foreach (Match match in anchorMatches)
                                {
                                    var content = match.Groups[2].Value;
                                    if (Regex.IsMatch(content, "unsubscribe", RegexOptions.IgnoreCase))
                                    {
                                        var hrefMatch = Regex.Match(match.Groups[1].Value, "href=\"(.*?)\"", RegexOptions.IgnoreCase);
                                        if (hrefMatch.Success)
                                        {
                                            group.UnsubscribeLink = hrefMatch.Groups[1].Value;
                                            group.UnsubscribeDate = item.ReceivedTime;
                                            break;
                                        }
                                    }

                                }

                            }

                            group.AddMailItem(item);
                        }
                        catch (Exception ex)
                        {
                            exceptionCount++;

                            Console.Error.WriteLine("Hiccup: " + ex.Message);
                            if (exceptionCount > 10) throw;
                        }

                    });

                    foreach (var group in groups.Values)
                    {
                        if (group.TotalCount < 5) continue;
                        //Console.WriteLine($"Group: {group.Domain} ({group.TotalCount})");
                        var examples = new List<string>();
                        foreach (var name in group.Addresses.Keys)
                        {
                            examples.Add(name);
                        }
                        Console.WriteLine($"{group.TotalCount}\t{group.Domain}\t{group.MostRecent}\t{group.Recipient}\t{String.Join(" | ", examples)}\t{group.UnsubscribeLink}");

                    }

                    if(deleteOlderMail)
                    {
                        Console.WriteLine("Deleting older items");
                        var deleteCount = 0;
                        scanner.ProcessFolder(folder, (item) =>
                        {
                            try
                            {
                                var domain = getDomainName(item.SenderEmailAddress);
                                if (!groups.ContainsKey(domain)) return;

                                var group = groups[domain];

                                if(group.TotalCount >= 5 && item.ReceivedTime < group.MostRecent)
                                {
                                    item.Delete();
                                    deleteCount++;
                                    if(deleteCount % 50 == 0)
                                    {
                                        Console.Write($"{deleteCount}        \r");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                exceptionCount++;

                                Console.Error.WriteLine("Hiccup: " + ex.Message);
                                if (exceptionCount > 10) throw;
                            }

                        });
                        Console.WriteLine();
                        Console.WriteLine("Done. Press ENTER to exit after outlook is done deleting stuff.");
                        Console.ReadLine();
                        
                    }                   
                }
                catch (System.Exception ex)
                {
                    Console.WriteLine("An error occurred: " + ex.ToString());
                }
            });
        }



        // --------------------------------------------------------------------------------
        // --------------------------------------------------------------------------------
        static void FarmContacts() {

            string normalizeName(string emailAddress, string name) {
                var originalName = name;
                if(string.IsNullOrEmpty(name?.Trim()) || name.Contains("@"))
                {
                    name = emailAddress.Split(new char[] { '@' })[0].Replace(".", "").Replace("-", "").Replace("_", "");
                }

                name = name.Trim().Trim('\'', '"');
                if(name.Contains(','))
                {
                    var parts = name.Split(new char[] { ',' }, 2);
                    name = parts[1].Trim() + " " + parts[0].Trim();
                }

                var fixedName = new StringBuilder();
                var nameParts = name.Split(' ');
                if(nameParts.Length == 1)
                {
                    fixedName.Append(nameParts[0]);
                }
                else
                {
                    foreach(var part in name.Split(' '))
                    {
                        var lowerPart = part.Trim().ToLower();
                        if (lowerPart == "" || Regex.IsMatch(lowerPart, "\\d")) continue;
                        if (fixedName.Length > 0) fixedName.Append(" ");
                        fixedName.Append(char.ToUpper(lowerPart[0]));
                        fixedName.Append(lowerPart.Substring(1));
                    }
                }

                var output = fixedName.ToString();

                if (originalName != output)
                {
                    Console.WriteLine($"Fixed Name:\n    old: {originalName}\n    new: {output}");
                }


                return output;
            }



            // --------------------------------------------------------------------------------
            RunInMailContext((outlookApp, mailRoot, outlookNamespace) =>
            {
                try
                {
                    Outlook.MAPIFolder sentItems = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);
                    Outlook.Items sentMailItems = sentItems.Items;

                    Outlook.MAPIFolder contactsFolder = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                    Outlook.Items contactItems = contactsFolder.Items;

                    var contactLookup = new Dictionary<string, Outlook.ContactItem>();
                    foreach (Outlook.ContactItem contact in contactItems)
                    {
                        foreach(string address in new string[] { contact.Email1Address, contact.Email2Address, contact.Email3Address })
                        {
                            if(!string.IsNullOrEmpty(address))
                            {
                                contactLookup[address.ToLower()] = contact;
                                Console.WriteLine($"CONTACT: {address}");
                            }

                        }
                    }
                    Console.WriteLine($"Found {contactLookup.Count} contacts.\n\n");

                    foreach (var item in sentMailItems)
                    {
                        if (!(item is Outlook.MailItem)) continue;
                        var mail = item as Outlook.MailItem;
                        var department = "";
                        var year = mail.SentOn.Year;
                        var dateId = mail.SentOn.ToString("yyyyMMdd");

                        var recipientText = new StringBuilder();
                        foreach (Outlook.Recipient r in mail.Recipients)
                        {
                            recipientText.Append(r.Address + " ");
                        }

                        var mailText = mail.Subject + " " + (mail.Body ?? "").Trim() + " " + recipientText.ToString();
                        mailText = mailText.ToLower();
                        if (Regex.IsMatch(mailText, "(villagetheatre|kidstage)")) {
                            department = "KidStage";
                        } else if (Regex.IsMatch(mailText, "(newsies|tarzan|poppins|cinderella|starcatcher|little mermaid|yctiwy|tlm|gphs theatre|gphs drama|gphstheatre|gphs theater|5thavenue)")) {
                            department = "Drama";
                        } else if(Regex.IsMatch(mailText, "(mit alumni|mit.edu|\\[drool\\]|\\[xi\\])")) { 
                            department = "TEP";
                        } else if(Regex.IsMatch(mailText, "(activity days|young women|bishopric|priest|stake |primary|seminary|church|lds.org|mormon|d&c|sunday school|ward )")) { 
                            department = "Church";
                        }
                        

                        foreach (Outlook.Recipient recipient in mail.Recipients)
                        {
                            var normalizedName = normalizeName(recipient.Address, recipient.Name);

                            bool changed = false;
                            if(recipient == null)
                            {
                                Console.WriteLine($"WEIRD: recipient is null on {mail.Subject}");
                                continue;
                            }
                            if(string.IsNullOrEmpty(recipient.Address))
                            {
                                Console.WriteLine($"WEIRD: address is null on {mail.Subject}");
                                continue;
                            }
                            contactLookup.TryGetValue(normalizedName, out var contact);
                            if(contact == null)
                            {
                                changed = true;
                                contact = (Outlook.ContactItem)outlookApp.CreateItem(Outlook.OlItemType.olContactItem);
                                contact.Email1Address = recipient.Address;
                                contact.FullName = normalizedName;
                                contact.Mileage = dateId;
                                contactLookup.Add(normalizedName, contact);
                            }

                            if(contact.Email1Address != recipient.Address && dateId.CompareTo(contact.Mileage) > 0)
                            {
                                contact.Mileage = dateId;
                                contact.Email1Address = recipient.Address;
                            }

                            if(department != "")
                            {
                                if(string.IsNullOrEmpty(contact.Department))
                                {
                                    contact.Department = department;
                                    changed = true;
                                } 
                                else if(contact.Department != department && contact.Department != "multi")
                                {
                                    contact.Department = "multi";
                                    changed = true;
                                }
                            }


                            Int32.TryParse(contact.BusinessFaxNumber, out int contactYear);
                            if(year > contactYear)
                            {
                                contact.BusinessFaxNumber = $"{year}";
                                changed = true;
                            }

                            if(string.IsNullOrEmpty(contact.FullName) && contact.FullName != normalizedName)
                            {
                                contact.FullName = normalizedName;
                                changed = true;
                            }
                            
                            if(changed)
                            {
                                Console.WriteLine($"\n{contact.Department}|{contact.BusinessFaxNumber}|{contact.Email1Address}|{contact.FullName}|{mail.SentOn}|{mail.Subject}");
                                //newContact.Save();
                            }
                            Console.Write(".");



                        }
                    }

                    Console.WriteLine("======================================================");
                    Console.WriteLine("======================================================");
                    Console.WriteLine("======================================================");
                    Console.WriteLine("======================================================");
                    Console.WriteLine("======================================================");

                    // gather up duplicate addresses
                    var addressLookup = new Dictionary<string, List<Outlook.ContactItem>>();
                    var output = new List<Outlook.ContactItem>();
                    foreach (var contact in contactLookup.Values)
                    {
                        var key = contact.Email1Address.ToLower();
                        if(!addressLookup.ContainsKey(key))
                        {
                            addressLookup[key] = new List<Outlook.ContactItem>();
                        }
                        addressLookup[key].Add(contact);
                    }

                    // For each address, pick the best one (the lastest that does not contain "parent")
                    foreach(var contactSet in addressLookup.Values)
                    {
                        var chosen = contactSet[0];
                        for(int i = 1; i < contactSet.Count; i++)
                        {
                            var item = contactSet[i];
                            if (item.FullName.ToLower().Contains("parent")) continue;
                            if(item.Mileage.CompareTo(chosen.Mileage) > 0)
                            {
                                chosen = item;
                            }
                        }
                        output.Add(chosen);
                    }

                    foreach (var contact in output)
                    {
                        Console.WriteLine($"{contact.Department}|{contact.BusinessFaxNumber}|{contact.Email1Address}|{contact.FullName}");
                    }

                    Console.WriteLine("DONE");
                }
                catch (System.Exception ex)
                {
                    Console.WriteLine("An error occurred: " + ex.ToString());
                }
            });
        }

        // --------------------------------------------------------------------------------
        // --------------------------------------------------------------------------------
        static void RunSpamWatch(string[] args)
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

                RunInMailContext((_mailApp, mailRoot, mapiNameSpace) =>
                {
                    while (true)
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
                        for (int i = cycleMinutes; i >= 1; i--)
                        {
                            Console.WriteLine($"Running again in {i} minutes ...");
                            Thread.Sleep(TimeSpan.FromMinutes(1));
                        }


                    }

                });

            }
            catch (System.Exception ex)
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
