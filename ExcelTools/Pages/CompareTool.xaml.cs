using System.Windows;
using System.Windows.Controls;
using ExcelTools.Attributes;

namespace ExcelTools.Pages
{
    [PageInfo(Header = "Compare Excels")]
    public partial class CompareTool : Page
    {
        private Logger Logger { get; }
        public CompareTool(Logger logger)
        {
            this.InitializeComponent();

            this.Logger = logger;
        }

        private void OnFileSelected()
        {
            this.CompareButton.IsEnabled = true;
        }

        private void OnFileChanged()
        {
            this.CompareButton.IsEnabled = false;
        }

        private void RunAnalysis(object sender, RoutedEventArgs e)
        {
            
        }
    }
}
