using System.IO;
using System.Windows;
using ExcelTools.Popups;
using Microsoft.Win32;

namespace ExcelTools
{
    public partial class MainWindow : Window
    {
        private Logger Logger { get; }
        private ExcelAnalysis ExcelAnalysis { get; }
        private ExcelWrapper ExcelWrapper { get; set; }
        public MainWindow()
        {
            this.InitializeComponent();

            this.Logger = new Logger(this.LogStackPanel, true);
            this.ExcelAnalysis = new ExcelAnalysis(this.Logger);
        }

        private void RunAnalysisHandler(object sender, RoutedEventArgs e)
        {
            if (this.ExcelWrapper == null)
            {
                return;
            }

            int index = this.RunAnalysisComboBox.SelectedIndex;

            if (index == 0)
            {
                var settings = new DuplicateAnalysis();
                settings.ShowDialog();

                this.ExcelAnalysis.FindDuplicates(this.ExcelWrapper, settings.ColumnInput.Text);
            }
        }

        private void ClearLogStackPanelHandler(object sender, RoutedEventArgs e)
        {
            this.LogStackPanel.Children.Clear();
        }

        private void LoadFileHandler(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog {Filter = "Excel Files|*.xls;*.xlsx|CSV files (*.csv)|*.csv" };

            if (openFileDialog.ShowDialog() == true)
            {
                this.ExcelWrapper = new ExcelWrapper(openFileDialog.FileName);
                this.LoadedFileNameTextbox.Text = Path.GetFileName(openFileDialog.FileName);
            }
        }
    }
}
