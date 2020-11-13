using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ExcelTools.Attributes;
using Microsoft.Win32;

namespace ExcelTools.Pages
{
    [PageInfo("Multi Counter")]
    public partial class MultiCounter : Page
    {
        private Logger Logger { get; }
        private ExcelAnalysis ExcelAnalysis { get; }
        private string[] FilePaths { get; set; }

        public MultiCounter(Logger logger)
        {
            this.InitializeComponent();

            this.Logger = logger;

            this.ExcelAnalysis = new ExcelAnalysis(this.Logger);
        }

        private void SelectFileHandler(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog { Filter = "Excel Files|*.xls;*.xlsx|CSV files (*.csv)|*.csv", Multiselect = true};

            if (openFileDialog.ShowDialog() == true)
            {
                this.FilePathTextBox.Text = string.Join(", ", openFileDialog.FileNames.Select(Path.GetFileName));
                this.FilePaths = openFileDialog.FileNames;

                this.SelectFileButton.Visibility = Visibility.Hidden;

                this.FilePathViewWrapper.Visibility = Visibility.Visible;
            }
        }

        private void ChangeFileHandler(object sender, RoutedEventArgs e)
        {
            this.FilePathTextBox.Clear();

            this.FilePathViewWrapper.Visibility = Visibility.Hidden;
            this.SelectFileButton.Visibility = Visibility.Visible;
        }

        private void RunAnalysis(object sender, RoutedEventArgs e)
        {
            this.ExcelAnalysis.MultipleFilesCountCells(this.FilePaths);
        }
    }
}