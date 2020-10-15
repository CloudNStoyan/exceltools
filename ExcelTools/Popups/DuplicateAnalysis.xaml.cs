using System.IO;
using System.Windows;
using Microsoft.Win32;

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

        private void LoadFileHandler(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog { Filter = "Excel Files|*.xls;*.xlsx|CSV files (*.csv)|*.csv" };

            if (openFileDialog.ShowDialog() == true)
            {
                this.LoadedFileNameTextbox.Text = Path.GetFileName(openFileDialog.FileName);
            }
        }
    }
}
