using System;

namespace ExcelTools.Attributes
{
    [AttributeUsage(AttributeTargets.Class)]
    public class PageInfo : Attribute
    {
        public string Header { get; set; }
        public int Order = 10;
        public bool ShowInNavigation = true;
    }
}
