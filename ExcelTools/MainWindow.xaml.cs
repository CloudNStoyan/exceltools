using System.Windows;
using ExcelTools.Pages;

namespace ExcelTools
{
    public partial class MainWindow : Window
    {
        private Logger Logger { get; }

        public MainWindow()
        {
            this.InitializeComponent();

            this.Logger = new Logger(this.LogStackPanel, true);

            var page = new FindDuplicates(this.Logger);
            this.Settings.Navigate(page);
        }

        private void ClearLogStackPanelHandler(object sender, RoutedEventArgs e)
        {
            this.LogStackPanel.Children.Clear();
        }

        private void DuplicateFinderHandler(object sender, RoutedEventArgs e)
        {
            var page = new FindDuplicates(this.Logger);
            this.Settings.Navigate(page);
        }

        private void MultiCounterHandler(object sender, RoutedEventArgs e)
        {
            var page = new MultiCounter(this.Logger);
            this.Settings.Navigate(page);
        }

        private void ColorCounterHandler(object sender, RoutedEventArgs e)
        {
            var page = new CountByColor(this.Logger);
            this.Settings.Navigate(page);
        }

        private void FindHandler(object sender, RoutedEventArgs e)
        {
            var page = new FindTool(this.Logger);
            this.Settings.Navigate(page);
        }

        private void ExportHandler(object sender, RoutedEventArgs e)
        {
            var page = new ExportTool(this.Logger);
            this.Settings.Navigate(page);
        }
    }
}