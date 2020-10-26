using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace ExcelTools.Pages
{
    public partial class CountByColor : Page
    {
        private Logger Logger { get; }
        private ExcelAnalysis ExcelAnalysis { get; }
        public CountByColor(Logger logger)
        {
            this.InitializeComponent();

            this.Logger = logger;

            this.ExcelAnalysis = new ExcelAnalysis(this.Logger);
        }

        private void SelectFileHandler(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog { Filter = "Excel Files|*.xls;*.xlsx|CSV files (*.csv)|*.csv" };

            if (openFileDialog.ShowDialog() == true)
            {
                this.FilePathTextBox.Text = openFileDialog.FileName;
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
            var excelWrapper = new ExcelWrapper(this.FilePathTextBox.Text);
            var a = excelWrapper.GetCountByColor("");

            foreach (string s in a)
            {
                this.Logger.Log(s);
            }
        }
    }
}
