using System.Windows;

namespace ExcelTools
{
    public partial class MainWindow : Window
    {
        private Logger Logger { get; }
        private ExcelAnalysis ExcelAnalysis { get; }
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
                this.ExcelAnalysis.FindDuplicates();
            }
        }

        private void ClearLogStackPanelHandler(object sender, RoutedEventArgs e)
        {
            this.LogStackPanel.Children.Clear();
        }
    }
}
