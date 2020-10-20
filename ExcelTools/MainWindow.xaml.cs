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
    }
}
