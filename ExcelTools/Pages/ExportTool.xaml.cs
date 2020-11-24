using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ExcelTools.Attributes;

namespace ExcelTools.Pages
{
    [PageInfo(Header = "Export Tool")]
    public partial class ExportTool : Page
    {
        private Logger Logger { get; }
        private ExcelAnalysis ExcelAnalysis { get; }
        private ExportSettings Settings { get; }

        public ExportTool(Logger logger)
        {
            this.InitializeComponent();

            this.Logger = logger;

            this.ExcelAnalysis = new ExcelAnalysis(this.Logger);

            this.Settings = new ExportSettings(this.SettingsContainer);
        }

        private void RunAnalysis(object sender, RoutedEventArgs e)
        {
            if (this.FileSelection.MultipleFilesChecked == false)
            {
                string filePath = this.FileSelection.SelectedFile;

                if (!File.Exists(filePath))
                {
                    MessageBox.Show("No file selected!");
                    return;
                }

                var excelWrapper = new ExcelWrapper(filePath);

                var data = this.ExcelAnalysis.ExportTool(excelWrapper, this.ColumnTextBox.Text, this.Settings.SkipEmpty);
                
                if (data == null)
                {
                    MessageBox.Show($"There is no column {this.ColumnTextBox.Text} in {excelWrapper.FileName}");
                    return;
                }

                data = data.Select(x => x?.Replace("\n", " ")).ToArray();
                this.OutputTextbox.Text = string.Join(this.Settings.SelectedSeparator, data);
            }
            else
            {
                string[] filePaths = this.FileSelection.SelectedFiles;

                if (filePaths == null || filePaths.Length == 0)
                {
                    MessageBox.Show("No file selected!");
                    return;
                }

                foreach (string selectedFile in filePaths)
                {
                    if (!File.Exists(selectedFile))
                    {
                        MessageBox.Show($"Cannot find {selectedFile}!");
                        return;
                    }
                }

                var excelWrappers = filePaths.Select(filePath => new ExcelWrapper(filePath)).ToArray();

                var data = this.ExcelAnalysis.ExportTool(excelWrappers, this.ColumnTextBox.Text, this.Settings.SkipEmpty);
                
                if (data == null)
                {
                    MessageBox.Show($"No data can be exported");
                    return;
                }
                
                data = data.Select(x => x?.Replace("\n", " ")).ToArray();
                this.OutputTextbox.Text = string.Join(this.Settings.SelectedSeparator, data);
            }
        }
    }

    public class ExportSettings
    {
        private StackPanel Container { get; }
        public bool SkipEmpty { get; set; }

        public ExportSettings(StackPanel container)
        {
            this.Container = container;

            var separationContainer = new StackPanel
            {
                Orientation = Orientation.Horizontal
            };

            var separationLabel = new Label
            {
                Content = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Children =
                    {
                        new TextBlock
                        {
                            Text = "Separator"
                        },
                        new TextBlock
                        {
                            Text = "*The symbol that will be used to split the data rows*",
                            FontSize = 10,
                            VerticalAlignment = VerticalAlignment.Bottom,
                            Margin = new Thickness(5,0,0,0),
                            Foreground = Brushes.Gray
                        }
                    }
                }
            };

            this.SetupSeparationSettings(separationContainer);

            var skipEmptyCheckBox = new CheckBox
            {
                Content = "Skip empty cells",
                Margin = new Thickness(5)
            };

            skipEmptyCheckBox.Click += (sender, args) => { this.SkipEmpty = skipEmptyCheckBox.IsChecked == true; };

            this.Container.Children.Add(separationLabel);
            this.Container.Children.Add(separationContainer);
            this.Container.Children.Add(skipEmptyCheckBox);
        }

        private readonly Dictionary<string, string> separators = new Dictionary<string, string>
        {
            { "Space", " "},
            { "New Line", "\r\n"},
            { "Comma", ", "}
        };

        public string SelectedSeparator { get; set; }

        private void SelectSeparator(object sender, RoutedEventArgs e)
        {
            var radioButton = (RadioButton)sender;

            this.SelectedSeparator = this.separators[radioButton.Content.ToString()];
        }

        private void SetupSeparationSettings(StackPanel separatorContainer)
        {
            var radioButtonsContainer = separatorContainer;

            foreach (var pair in this.separators)
            {
                var radioButton = new RadioButton
                {
                    GroupName = "Separators",
                    Content = pair.Key,
                    Margin = new Thickness(40, 0, 0, 0),
                };

                radioButton.Click += this.SelectSeparator;

                radioButtonsContainer.Children.Add(radioButton);
            }
        }
    }
}