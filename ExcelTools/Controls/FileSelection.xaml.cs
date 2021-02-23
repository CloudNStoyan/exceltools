using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;

namespace ExcelTools.Controls
{
    public partial class FileSelection : UserControl
    {
        private const string File = "File";
        private const string Files = "Files";
        private const string SelectFile = "Select File";
        private const string SelectFiles = "Select Files";
        private const string ExcelFileAnalysis = "*The excel file you want to analyse*";
        private const string ExcelFilesAnalysis = "*The excel files you want to analyse*";
        public string HeaderText
        {
            get => (string)this.GetValue(HeaderTextProperty);
            set => this.SetValue(HeaderTextProperty, value);
        }

        public static readonly DependencyProperty HeaderTextProperty
            = DependencyProperty.Register(
                nameof(HeaderText),
                typeof(string),
                typeof(FileSelection),
                new PropertyMetadata(File)
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
                nameof(Selection),
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
                nameof(SubHeaderText),
                typeof(string),
                typeof(FileSelection),
                new PropertyMetadata(ExcelFileAnalysis)
            );

        public bool MultipleFilesChecked
        {
            get => (bool)this.GetValue(MultipleFilesCheckedProperty);
            set => this.SetValue(MultipleFilesCheckedProperty, value);
        }

        public static readonly DependencyProperty MultipleFilesCheckedProperty
            = DependencyProperty.Register(
                nameof(MultipleFilesChecked),
                typeof(bool),
                typeof(FileSelection),
                new PropertyMetadata(false)
            );

        public bool FileIsSelected
        {
            get => (bool)this.GetValue(FileIsSelectedProperty);
            set => this.SetValue(FileIsSelectedProperty, value);
        }

        public static readonly DependencyProperty FileIsSelectedProperty
            = DependencyProperty.Register(
                nameof(FileIsSelected),
                typeof(bool),
                typeof(FileSelection),
                new PropertyMetadata(false)
            );

        public event Action FileSelected;
        public event Action FileChanged;

        public string SelectedFile { get; set; }
        public string[] SelectedFiles { get; set; }
        private SavedData.SavedData SavedData { get; set; }

        public FileSelection()
        {
            this.InitializeComponent();

            this.SavedData = new SavedData.SavedData();
        }

        private void SelectFileHandler(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog { Filter = "Excel Files|*.xls;*.xlsx|CSV files (*.csv)|*.csv" };

            if (this.Selection == SelectionType.Single)
            {
                if (openFileDialog.ShowDialog() != true)
                {
                    return;
                }

                string fileName = openFileDialog.FileName;

                this.FilePathTextBox.Text = fileName;

                this.SelectedFile = fileName;
                this.FileIsSelected = true;

                this.SavedData.Config.RecentFiles.Add(openFileDialog.FileName);
                this.SavedData.Save();

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
                this.FileIsSelected = true;

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

                    this.SavedData.Config.RecentFiles.Add(openFileDialog.FileName);
                    this.SavedData.Save();
                }

                this.FileIsSelected = true;

                this.SelectFileButton.Visibility = Visibility.Hidden;

                this.FilePathViewWrapper.Visibility = Visibility.Visible;
            }

            this.FileSelected?.Invoke();
        }

        public void ChangeFileHandler(object sender, RoutedEventArgs e)
        {
            this.FilePathTextBox.Clear();
            this.FileIsSelected = false;

            this.FilePathViewWrapper.Visibility = Visibility.Hidden;
            this.SelectFileButton.Visibility = Visibility.Visible;

            this.FileChanged?.Invoke();
        }

        private void MultipleFiles_OnClick(object sender, RoutedEventArgs e)
        {
            if (this.MultipleFiles.IsChecked == true)
            {
                this.SelectFileButton.Content = SelectFiles;
                this.LabelFileText.Header = Files;
                this.LabelFileText.SubHeader = ExcelFilesAnalysis;

                this.MultipleFilesChecked = true;
            }
            else
            {
                this.SelectFileButton.Content = SelectFile;
                this.LabelFileText.Header = File;
                this.LabelFileText.SubHeader = ExcelFileAnalysis;

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
                this.SelectFileButton.Content = SelectFile;
                this.LabelFileText.Header = File;
                this.LabelFileText.SubHeader = ExcelFileAnalysis;
            }
            else if (this.Selection == SelectionType.Multi)
            {
                this.MultipleFilesLabel.Visibility = Visibility.Collapsed;
                this.SelectFileButton.Content = SelectFiles;
                this.LabelFileText.Header = Files;
                this.LabelFileText.SubHeader = ExcelFilesAnalysis;
            }
            else if (this.Selection == SelectionType.Both)
            {
                this.MultipleFilesLabel.Visibility = Visibility.Visible;
                this.SelectFileButton.Content = SelectFile;
                this.LabelFileText.Header = File;
                this.LabelFileText.SubHeader = ExcelFileAnalysis;
            }
        }

        private void LoadContextMenu()
        {
            string[] recentFiles = this.SavedData.Config?.RecentFiles?.ToArray();

            if (recentFiles != null && recentFiles.Length > 0)
            {
                var contextMenu = new ContextMenu();

                foreach (string recentFile in recentFiles)
                {
                    var menuItem = new MenuItem { Header = recentFile };

                    menuItem.Click += (o, args) =>
                    {
                        if (this.Selection == SelectionType.Single)
                        {
                            this.SelectedFile = recentFile;

                            this.FilePathTextBox.Text = recentFile;
                        }
                        else if (this.Selection == SelectionType.Multi)
                        {
                            this.SelectedFiles = new[] {recentFile};

                            this.FilePathTextBox.Text = recentFile;
                        }
                        else if (this.Selection == SelectionType.Both)
                        {
                            if (this.MultipleFilesChecked)
                            {
                                this.SelectedFiles = new[] { recentFile };

                                this.FilePathTextBox.Text = recentFile;
                            }
                            else
                            {
                                this.SelectedFile = recentFile;

                                this.FilePathTextBox.Text = recentFile;
                            }
                        }

                        
                        this.FileIsSelected = true;

                        this.SelectFileButton.Visibility = Visibility.Hidden;

                        this.FilePathViewWrapper.Visibility = Visibility.Visible;

                        this.FileSelected?.Invoke();
                    };

                    contextMenu.Items.Add(menuItem);
                }

                this.SelectFileButton.ContextMenu = contextMenu;
            }
        }

        private void SelectFileButton_OnLoaded(object sender, RoutedEventArgs e) => this.LoadContextMenu();
        private void SelectFileButton_OnMouseRightButtonDown(object sender, MouseButtonEventArgs e) =>
            this.LoadContextMenu();
    }
}
