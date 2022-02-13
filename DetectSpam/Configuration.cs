using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DetectSpam
{
    public class MoveListItem
    {
        public string HeaderText { get; set; }
        public string TargetFolder { get; set; }
    }

    public class WordScore
    {
        public string RegExp { get; set; }
        public int Score { get; set; }
    }

    public class Configuration
    {
        public string[] SpamFolderPaths { get; set; }
        public int OKCutoffScore { get; set; }
        public string DefaultOutputFolder  { get; set;}
        public MoveListItem[] MoveRules { get; set; }
        public WordScore[] WordScoreRules { get; set; }
        public string[] WhiteListTextPatterns { get; set; }
        public string[] WhiteListHtmlPatterns { get; set; }
    }
}
