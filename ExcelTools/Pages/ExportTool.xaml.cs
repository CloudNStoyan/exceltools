using System;
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

        public ExportTool(Logger logger)
        {
            this.InitializeComponent();

            this.Logger = logger;
        }

        private string[] separators = {" ", Environment.NewLine, ","};

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

                var data = this.Export(excelWrapper, this.ColumnTextBox.Text, this.SkipEmpty.IsChecked == true);
                
                if (data == null)
                {
                    MessageBox.Show($"There is no column {this.ColumnTextBox.Text} in {excelWrapper.FileName}");
                    return;
                }

                data = data.Select(x => x?.Replace("\n", " ")).ToArray();

                int separatorIndex = int.Parse(this.SeparatorsPanel.Children.OfType<RadioButton>()
                    .FirstOrDefault(r => r.IsChecked == true)
                    ?.DataContext.ToString());

                string separator = this.separators[separatorIndex];

                this.OutputTextbox.Text = string.Join(separator, data);
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

                var data = this.Export(excelWrappers, this.ColumnTextBox.Text, this.SkipEmpty.IsChecked == true);
                
                if (data == null)
                {
                    MessageBox.Show($"No data can be exported");
                    return;
                }
                
                data = data.Select(x => x?.Replace("\n", " ")).ToArray();

                int separatorIndex = int.Parse(this.SeparatorsPanel.Children.OfType<RadioButton>()
                    .FirstOrDefault(r => r.IsChecked == true)
                    ?.DataContext.ToString());

                string separator = this.separators[separatorIndex];

                this.OutputTextbox.Text = string.Join(separator, data);
            }
        }

        private string[] Export(ExcelWrapper excelWrapper, string column, bool skipEmpty = false)
        {
            int columnNumber = ExcelWrapper.ConvertStringColumnToNumber(column);

            if (columnNumber == -1)
            {
                MessageBox.Show($"Column '{column}' is not a valid column!");
                return null;
            }

            string[] columnData = !skipEmpty ? excelWrapper.GetStringRows(columnNumber) : excelWrapper.GetStringRows(columnNumber)?.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();

            return columnData;
        }

        private string[] Export(ExcelWrapper[] excelWrappers, string column, bool skipEmpty = false)
        {
            int columnNumber = ExcelWrapper.ConvertStringColumnToNumber(column);

            if (columnNumber == -1)
            {
                MessageBox.Show($"Column '{column}' is not a valid column!");
                return null;
            }

            var list = new List<string>();

            foreach (var excelWrapper in excelWrappers)
            {
                string[] columnData = !skipEmpty ? excelWrapper.GetStringRows(columnNumber) : excelWrapper.GetStringRows(columnNumber)?.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();

                if (columnData == null)
                {
                    MessageBox.Show($"{excelWrapper.FileName} doesn't have '{column}' column!");
                    return null;
                }

                list.AddRange(columnData);
            }

            return list.ToArray();
        }
    }
}