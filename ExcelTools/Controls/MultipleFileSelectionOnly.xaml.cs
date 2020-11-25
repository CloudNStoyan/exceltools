using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace ExcelTools.Controls
{
    public partial class MultipleFileSelectionOnly : UserControl
    {
        public string[] SelectedFiles { get; private set; }

        public event Action FileSelected;
        public event Action FileChanged;

        public MultipleFileSelectionOnly()
        {
            this.InitializeComponent();
        }

        private void SelectFileHandler(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = CustomResources.ExcelFileFilter,
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() != true)
            {
                return;
            }

            string[] filePaths = openFileDialog.FileNames;

            this.FilePathTextBox.Text = string.Join(",", filePaths.Select(Path.GetFileName));
            this.SelectedFiles = filePaths;

            this.SelectFileButton.Visibility = Visibility.Hidden;

            this.FilePathViewWrapper.Visibility = Visibility.Visible;

            this.FileSelected?.Invoke();
        }

        private void ChangeFileHandler(object sender, RoutedEventArgs e)
        {
            this.FilePathTextBox.Clear();

            this.FilePathViewWrapper.Visibility = Visibility.Hidden;
            this.SelectFileButton.Visibility = Visibility.Visible;

            this.FileChanged?.Invoke();
        }
    }
}