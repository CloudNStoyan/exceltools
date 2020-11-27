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

        private void FileSelection_OnLoaded(object sender, RoutedEventArgs e)
        {
            this.DataContext = this;
        }
    }
}
