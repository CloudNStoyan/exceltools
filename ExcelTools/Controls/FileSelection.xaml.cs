using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace ExcelTools.Controls
{
    public partial class FileSelection : UserControl
    {
        public string SelectedFile { get; set; }

        public FileSelection()
        {
            this.InitializeComponent();
        }

        private void SelectFileHandler(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog { Filter = CustomResources.ExcelFileFilter };

            if (openFileDialog.ShowDialog() != true)
            {
                return;
            }

            string fileName = openFileDialog.FileName;

            this.FilePathTextBox.Text = fileName;

            this.SelectedFile = fileName;

            this.SelectFileButton.Visibility = Visibility.Hidden;

            this.FilePathViewWrapper.Visibility = Visibility.Visible;
        }

        private void ChangeFileHandler(object sender, RoutedEventArgs e)
        {
            this.FilePathTextBox.Clear();

            this.FilePathViewWrapper.Visibility = Visibility.Hidden;
            this.SelectFileButton.Visibility = Visibility.Visible;
        }
    }
}
