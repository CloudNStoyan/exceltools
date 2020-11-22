using System.IO;
using System.Windows;
using System.Windows.Controls;
using ExcelTools.Attributes;

namespace ExcelTools.Pages
{
    [PageInfo(Header = "Duplicate Finder", Order = 0)]
    public partial class FindDuplicates : Page
    {
        private Logger Logger { get; }
        private ExcelAnalysis ExcelAnalysis { get; }
        public FindDuplicates(Logger logger)
        {
            this.InitializeComponent();

            this.Logger = logger;

            this.ExcelAnalysis = new ExcelAnalysis(this.Logger);
        }

        private void RunAnalysis(object sender, RoutedEventArgs e)
        {
            if (!File.Exists(this.FileSelection.SelectedFile))
            {
                MessageBox.Show("No file selected!");
                return;
            }

            var excelWrapper = new ExcelWrapper(this.FileSelection.SelectedFile);
            this.ExcelAnalysis.FindDuplicates(excelWrapper, this.ColumnTextBox.Text);
        }
    }
}