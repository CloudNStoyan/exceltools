using System.Windows;

namespace ExcelTools.Popups
{
    /// <summary>
    /// Interaction logic for DuplicateAnalysis.xaml
    /// </summary>
    public partial class DuplicateAnalysis : Window
    {
        public DuplicateAnalysis()
        {
            this.InitializeComponent();
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
