using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ExcelTools.Alerts;
using ExcelTools.Attributes;

namespace ExcelTools.Pages
{
    [PageInfo(Header = "Combine Tool")]
    public partial class CombineTool : Page
    {
        public CombineTool()
        {
            this.InitializeComponent();
        }

        private void FileSelection_OnFileSelected()
        {
            string[] files = this.FileSelection.SelectedFiles;

            if (files.Length >= 2)
            {
                var excelWrappers = files.Select(x => new ExcelWrapper(x));



            }
            else
            {
                AlertManager.Custom("You need at least 2 files to combine them");
            }
        }
    }
}
