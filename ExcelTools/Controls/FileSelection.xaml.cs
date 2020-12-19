using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace ExcelTools.Controls
{
    public partial class FileSelection : UserControl
    {
        public string HeaderText
        {
            get => (string)this.GetValue(HeaderTextProperty);
            set => this.SetValue(HeaderTextProperty, value);
        }

        public static readonly DependencyProperty HeaderTextProperty
            = DependencyProperty.Register(
                "HeaderText",
                typeof(string),
                typeof(FileSelection),
                new PropertyMetadata("File")
            );

        public enum SelectionType
        {
            Single,
            Multi,
            Both
        }

        public SelectionType Selection
        {
            get => (SelectionType)this.GetValue(SelectionProperty);
            set => this.SetValue(SelectionProperty, value);
        }

        public static readonly DependencyProperty SelectionProperty
            = DependencyProperty.Register(
                "Selection",
                typeof(SelectionType),
                typeof(FileSelection),
                new PropertyMetadata(SelectionType.Single)
            );

        public string SubHeaderText
        {
            get => (string)this.GetValue(SubHeaderTextProperty);
            set => this.SetValue(SubHeaderTextProperty, value);
        }

        public static readonly DependencyProperty SubHeaderTextProperty
            = DependencyProperty.Register(
                "SubHeaderText",
                typeof(string),
                typeof(FileSelection),
                new PropertyMetadata("*The excel file you want to analyse*")
            );

        public bool MultipleFilesChecked
        {
            get => (bool)this.GetValue(MultipleFilesCheckedProperty);
            set => this.SetValue(MultipleFilesCheckedProperty, value);
        }

        public static readonly DependencyProperty MultipleFilesCheckedProperty
            = DependencyProperty.Register(
                "MultipleFilesChecked",
                typeof(bool),
                typeof(FileSelection),
                new PropertyMetadata(false)
            );

        public event Action FileSelected;
        public event Action FileChanged;

        public string SelectedFile { get; set; }
        public string[] SelectedFiles { get; set; }

        public FileSelection()
        {
            this.InitializeComponent();
        }

        private void SelectFileHandler(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog { Filter = CustomResources.ExcelFileFilter };

            if (this.Selection == SelectionType.Single)
            {
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
            else if (this.Selection == SelectionType.Multi)
            {
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() != true)
                {
                    return;
                }

                string[] filePaths = openFileDialog.FileNames;

                this.FilePathTextBox.Text = string.Join(",", filePaths.Select(Path.GetFileName));
                this.SelectedFiles = filePaths;

                this.SelectFileButton.Visibility = Visibility.Hidden;

                this.FilePathViewWrapper.Visibility = Visibility.Visible;
            }
            else if (this.Selection == SelectionType.Both)
            {
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

            this.FileSelected?.Invoke();
        }

        private void ChangeFileHandler(object sender, RoutedEventArgs e)
        {
            this.FilePathTextBox.Clear();

            this.FilePathViewWrapper.Visibility = Visibility.Hidden;
            this.SelectFileButton.Visibility = Visibility.Visible;

            this.FileChanged?.Invoke();
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

        private void FileSelection_OnLoaded(object sender, RoutedEventArgs e)
        {
            this.DataContext = this;

            if (this.Selection == SelectionType.Single)
            {
                this.MultipleFilesLabel.Visibility = Visibility.Collapsed;
            }
            else if (this.Selection == SelectionType.Multi)
            {
                this.MultipleFilesLabel.Visibility = Visibility.Collapsed;
                this.SelectFileButton.Content = "Select Files";
                this.FileTextBlock.Text = "Files";
                this.FileSubTextBlock.Text = "*The excel files you want to analyse*";

            }
            else if (this.Selection == SelectionType.Both)
            {
                this.MultipleFilesLabel.Visibility = Visibility.Visible;
            }
        }
    }
}
