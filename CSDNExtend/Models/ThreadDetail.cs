using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CSDNExtend.Models
{
    public class ThreadDetail
    {
       
        public string Team { get; set; }
        public string IsAnswered { get; set; }
        public string Owner { get; set; }
        [IsURLAttribute]
        public string URL { get; set; }
        [IsTitleAttribute]
        public string Title { get; set; }

        
        public string TechCategory { get; set; }
        public string IssueType { get; set; }
        public string IR { get; set; }
        public DateTime CreateOn { get; set; }
        public string FirstReply { get; set; }
        public string Labor { get; set; }
        public string Replies { get; set; }
        public string CssAction { get; set; }
        public string Replied { get; set; }
        public string Difficulty { get; set; }
        public string CustomLooking { get; set; }
        public string DayToAnswer { get; set; }
        public string Contribution { get; set; }
    }
    public class IsURLAttribute : Attribute
    {

    }
}