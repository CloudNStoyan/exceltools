using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace ExcelTools.Controls
{
    public partial class MultipleFileSelection : UserControl
    {
        public string SelectedFile { get; set; }
        public string[] SelectedFiles { get; set; }
        public bool MultipleFilesChecked { get; set; }

        public MultipleFileSelection()
        {
            this.InitializeComponent();
        }

        private void SelectFileHandler(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog { Filter = CustomResources.ExcelFileFilter };

            if (this.MultipleFiles.IsChecked == true)
            {
                openFileDialog.Multiselect = true;
            }

            if (openFileDialog.ShowDialog() != true)
            {
                return;
            }

            this.FilePathTextBox.Text = this.MultipleFiles.IsChecked == true
                ? string.Join(",", openFileDialog.FileNames.Select(Path.GetFileName))
                : openFileDialog.FileName;

            if (this.MultipleFiles.IsChecked == true)
            {
                this.SelectedFiles = openFileDialog.FileNames;
            }
            else
            {
                this.SelectedFile = openFileDialog.FileName;
            }

            this.SelectFileButton.Visibility = Visibility.Hidden;

            this.FilePathViewWrapper.Visibility = Visibility.Visible;
        }

        private void ChangeFileHandler(object sender, RoutedEventArgs e)
        {
            this.FilePathTextBox.Clear();

            this.FilePathViewWrapper.Visibility = Visibility.Hidden;
            this.SelectFileButton.Visibility = Visibility.Visible;
        }

        private void MultipleFiles_OnClick(object sender, RoutedEventArgs e)
        {
            if (this.MultipleFiles.IsChecked == true)
            {
                this.SelectFileButton.Content = "Select Files";
                this.FileTextBlock.Text = "Files";
                this.FileSubTextBlock.Text = "*The excel files you want to analyse*";

                this.MultipleFilesChecked = true;
            }
            else
            {
                this.SelectFileButton.Content = "Select File";
                this.FileTextBlock.Text = "File";
                this.FileSubTextBlock.Text = "*The excel file you want to analyse*";

                this.MultipleFilesChecked = false;
            }

            this.ChangeFileHandler(sender, e);
        }
    }
}
