using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ExcelTools.DataSaving;
using Microsoft.Win32;

namespace ExcelTools.Controls
{
    public partial class FileSelection : UserControl
    {
        private const string FileLabel = "File";
        private const string FilesLabel = "Files";
        private const string SelectFileLabel = "Select File";
        private const string SelectFilesLabel = "Select Files";
        private const string ExcelFileAnalysisLabel = "*The excel file you want to analyse*";
        private const string ExcelFilesAnalysisLabel = "*The excel files you want to analyse*";
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
                new PropertyMetadata(FileLabel)
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

        public bool Persistent
        {
            get => (bool) this.GetValue(PersistentProperty);
            set => this.SetValue(PersistentProperty, value);
        }

        public static readonly DependencyProperty PersistentProperty = 
            DependencyProperty.Register(
                nameof(Persistent),
                typeof(bool),
                typeof(FileSelection), 
                new PropertyMetadata(true)
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
                new PropertyMetadata(ExcelFileAnalysisLabel)
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

        public FileSelection()
        {
            this.InitializeComponent();
        }

        private void SelectFileHandler(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog { Filter = "Excel FilesLabel|*.xls;*.xlsx|CSV files (*.csv)|*.csv" };

            if (this.Selection == SelectionType.Single)
            {
                if (openFileDialog.ShowDialog() != true)
                {
                    return;
                }

                this.SelectFile(openFileDialog.FileName);
            }
            else if (this.Selection == SelectionType.Multi)
            {
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() != true)
                {
                    return;
                }

                this.SelectFile(openFileDialog.FileNames);
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

                if (this.MultipleFiles.IsChecked == true)
                {
                    this.SelectFile(openFileDialog.FileNames);
                }
                else
                {
                    this.SelectFile(openFileDialog.FileName);
                }
            }
        }

        private void SelectFile(string filepath) => this.SelectFile(new[] {filepath});

        private void SelectFile(string[] filepaths)
        {
            string[] fileNames = filepaths.Select(Path.GetFileName).ToArray();

            if (this.Selection == SelectionType.Multi || (this.Selection == SelectionType.Both && this.MultipleFilesChecked))
            {
                this.SelectedFiles = filepaths;
                this.FilePathTextBox.Text = string.Join(",", fileNames);
                SavedData.Config.AddToRecentFiles(filepaths);
            }
            else
            {
                this.SelectedFile = filepaths[0];
                this.FilePathTextBox.Text = fileNames[0];
                SavedData.Config.AddToRecentFiles(filepaths[0]);
            }

            SavedData.Save();

            this.FileIsSelected = true;

            this.SelectFileButton.Visibility = Visibility.Hidden;

            this.FilePathViewWrapper.Visibility = Visibility.Visible;

            this.FileSelected?.Invoke();

            if (!this.Persistent)
            {
                this.ChangeFileHandler(null, null);
            }
        }

        private void ChangeFileHandler(object sender, RoutedEventArgs e)
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
                this.SelectFileButton.Content = SelectFilesLabel;
                this.LabelFileText.Header = FilesLabel;
                this.LabelFileText.SubHeader = ExcelFilesAnalysisLabel;

                this.MultipleFilesChecked = true;
            }
            else
            {
                this.SelectFileButton.Content = SelectFileLabel;
                this.LabelFileText.Header = FileLabel;
                this.LabelFileText.SubHeader = ExcelFileAnalysisLabel;

                this.MultipleFilesChecked = false;
            }

            this.SelectFileButton.ContextMenu = null;
            this.ChangeFileHandler(sender, e);
        }

        private void FileSelection_OnLoaded(object sender, RoutedEventArgs e)
        {
            this.DataContext = this;

            if (this.Selection == SelectionType.Single)
            {
                this.MultipleFilesLabel.Visibility = Visibility.Collapsed;
                this.SelectFileButton.Content = SelectFileLabel;
                this.LabelFileText.Header = FileLabel;
                this.LabelFileText.SubHeader = ExcelFileAnalysisLabel;
            }
            else if (this.Selection == SelectionType.Multi)
            {
                this.MultipleFilesLabel.Visibility = Visibility.Collapsed;
                this.SelectFileButton.Content = SelectFilesLabel;
                this.LabelFileText.Header = FilesLabel;
                this.LabelFileText.SubHeader = ExcelFilesAnalysisLabel;
            }
            else if (this.Selection == SelectionType.Both)
            {
                this.MultipleFilesLabel.Visibility = Visibility.Visible;
                this.SelectFileButton.Content = SelectFileLabel;
                this.LabelFileText.Header = FileLabel;
                this.LabelFileText.SubHeader = ExcelFileAnalysisLabel;
            }
        }

        private void LoadContextMenu()
        {
            string[] recentFiles;

            bool isMulti = false;

            if (this.Selection == SelectionType.Multi || (this.Selection == SelectionType.Both && this.MultipleFilesChecked))
            {
                recentFiles = SavedData.Config.RecentMultipleFiles
                    .Select(x => string.Join(", ", x)).ToArray();
                isMulti = true;
            }
            else
            {
                recentFiles = SavedData.Config.RecentFiles.ToArray();
            }

            if (recentFiles.Length <= 0)
            {
                return;
            }

            var contextMenu = new ContextMenu();

            foreach (string recentFile in recentFiles)
            {
                var menuItem = new MenuItem {Header = recentFile};

                if (isMulti)
                {
                    menuItem.Header = string.Join(", ",
                        recentFile.Split(new[] {", "}, StringSplitOptions.None).Select(Path.GetFileName));
                }

                menuItem.Click += (o, args) =>
                {
                    this.SelectFile(recentFile);
                };

                contextMenu.Items.Add(menuItem);
            }

            this.SelectFileButton.ContextMenu = contextMenu;
        }

        private void SelectFileButton_OnMouseRightButtonDown(object sender, MouseButtonEventArgs e) =>
            this.LoadContextMenu();
    }
}
