using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace ExcelTools.Pages
{
    public partial class FindDuplicates : Page
    {
        private Logger Logger { get; }
        private ExcelAnalysis ExcelAnalysis { get; }
        public static string Header = "Duplicate Finder";
        public FindDuplicates(Logger logger)
        {
            this.InitializeComponent();

            this.Logger = logger;

            this.ExcelAnalysis = new ExcelAnalysis(this.Logger);
        }

        private void SelectFileHandler(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog { Filter = "Excel Files|*.xls;*.xlsx|CSV files (*.csv)|*.csv" };

            if (this.MultipleFiles.IsChecked == true)
            {
                openFileDialog.Multiselect = true;
            }

            if (openFileDialog.ShowDialog() != true)
            {
                return;
            }

            this.FilePathTextBox.Text = this.MultipleFiles.IsChecked == true
                ? string.Join(",", openFileDialog.FileNames.Select(Path.GetFileName))
                : openFileDialog.FileName;

            this.SelectFileButton.Visibility = Visibility.Hidden;

            this.FilePathViewWrapper.Visibility = Visibility.Visible;
        }

        private void ChangeFileHandler(object sender, RoutedEventArgs e)
        {
            this.FilePathTextBox.Clear();

            this.FilePathViewWrapper.Visibility = Visibility.Hidden;
            this.SelectFileButton.Visibility = Visibility.Visible;
        }

        private void RunAnalysis(object sender, RoutedEventArgs e)
        {
            if (this.MultipleFiles.IsChecked == false)
            {

                var excelWrapper = new ExcelWrapper(this.FilePathTextBox.Text);
                this.ExcelAnalysis.FindDuplicates(excelWrapper, this.ColumnTextBox.Text);
            }
        }

        private void MultipleFiles_OnClick(object sender, RoutedEventArgs e)
        {
            if (this.MultipleFiles.IsChecked == true)
            {
                this.SelectFileButton.Content = "Select Files";
                this.FileTextBlock.Text = "Files";
                this.FileSubTextBlock.Text = "*The excel files you want to analyse*";
            }
            else
            {
                this.SelectFileButton.Content = "Select File";
                this.FileTextBlock.Text = "File";
                this.FileSubTextBlock.Text = "*The excel file you want to analyse*";
            }

            this.ChangeFileHandler(sender, e);
        }
    }
}
