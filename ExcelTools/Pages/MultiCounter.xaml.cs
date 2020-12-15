using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ExcelTools.Attributes;

namespace ExcelTools.Pages
{
    [PageInfo(Header = "Multi Counter")]
    public partial class MultiCounter : Page
    {
        private Logger Logger { get; }

        public MultiCounter(Logger logger)
        {
            this.InitializeComponent();

            this.Logger = logger;
        }

        private void RunAnalysis(object sender, RoutedEventArgs e)
        {
            if (this.FileSelection.MultipleFilesChecked)
            {
                var filePaths = this.FileSelection.SelectedFiles;
            
                if (filePaths == null || filePaths.Length == 0)
                {
                    MessageBox.Show("No files selected!");
                    return;
                }
            
                foreach (string filePath in filePaths)
                {
                    if (!File.Exists(filePath))
                    {
                        MessageBox.Show($"Can't find '{filePath}'");
                        return;
                    }
                }

                int count = filePaths.Select(excelFile => new ExcelWrapper(excelFile)).Select(excelWrapper => excelWrapper.GetCount()).Sum();

                this.Logger.Log(count + " entries!");
            }
            else
            {
                string filePath = this.FileSelection.SelectedFile;
            
                if (!File.Exists(filePath))
                {
                    MessageBox.Show("No file selected!");
                    return;
                }

                var excelWrapper = new ExcelWrapper(filePath);

                this.Logger.Log(excelWrapper.GetCount().ToString());
            }
        }
    }
}