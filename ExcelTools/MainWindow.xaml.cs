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
            int index = this.RunAnalysisComboBox.SelectedIndex;

            if (index == 0)
            {
                var duplicateAnalysis = new DuplicateAnalysis();
                duplicateAnalysis.ShowDialog();

                this.ExcelAnalysis.FindDuplicates(this.ExcelWrapper, duplicateAnalysis.ColumnInput.Text);
            }
        }

        private void ClearLogStackPanelHandler(object sender, RoutedEventArgs e)
        {
            this.LogStackPanel.Children.Clear();
        }
    }
}
