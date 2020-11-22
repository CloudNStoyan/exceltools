using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ExcelTools.Attributes;

namespace ExcelTools.Pages
{
    [PageInfo(Header = "Find Tool")]
    public partial class FindTool : Page
    {
        private Logger Logger { get; }
        private ExcelAnalysis ExcelAnalysis { get; }

        public FindTool(Logger logger)
        {
            this.InitializeComponent();

            this.Logger = logger;

            this.ExcelAnalysis = new ExcelAnalysis(this.Logger);
        }

        private void RunAnalysis(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(this.FindValueInput.Text))
            {
                MessageBox.Show("Value cannot be empty or white space only!");
                return;
            }

            if (this.FileSelection.MultipleFilesChecked == false)
            {
                string filePath = this.FileSelection.SelectedFile;

                if (!File.Exists(filePath))
                {
                    MessageBox.Show("No file selected!");
                    return;
                }

                var excelWrapper = new ExcelWrapper(filePath);
                this.ExcelAnalysis.FindTool(excelWrapper, this.ColumnTextBox.Text, this.FindValueInput.Text, this.CaseSensitiveCheck.IsChecked == true);
            }
            else
            {
                string[] filePaths = this.FileSelection.SelectedFiles;

                if (filePaths == null || filePaths.Length == 0)
                {
                    MessageBox.Show("No file selected!");
                    return;
                }

                foreach (string selectedFile in filePaths)
                {
                    if (!File.Exists(selectedFile))
                    {
                        MessageBox.Show($"Cannot find {selectedFile}!");
                        return;
                    }
                }

                var excelWrappers = filePaths.Select(filePath => new ExcelWrapper(filePath)).ToList();

                this.ExcelAnalysis.FindTool(excelWrappers.ToArray(), this.ColumnTextBox.Text, this.FindValueInput.Text, this.CaseSensitiveCheck.IsChecked == true);
            }
        }
    }
}
