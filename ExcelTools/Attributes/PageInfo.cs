using System;

namespace ExcelTools.Attributes
{
    [AttributeUsage(AttributeTargets.Class)]
    public class PageInfo : Attribute
    {
        public string Header { get; set; }
        public PageInfo(string header)
        {
            this.Header = header;
        }
    }
}
