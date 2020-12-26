using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using ExcelTools.Attributes;

namespace ExcelTools.Pages
{
    [PageInfo(Header = "Find Tool")]
    public partial class FindTool : Page
    {
        private Logger Logger { get; }

        public FindTool(Logger logger)
        {
            this.InitializeComponent();

            this.Logger = logger;
        }

        private void RunAnalysis(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(this.FindValueInput.Text))
            {
                MessageBox.Show("Value cannot be empty or white space only!");
                return;
            }

            if (this.FileSelection.MultipleFilesChecked == false)
            {
                string filePath = this.FileSelection.SelectedFile;

                if (!File.Exists(filePath))
                {
                    MessageBox.Show("No file selected!");
                    return;
                }

                var excelWrapper = new ExcelWrapper(filePath);
                this.FindAnalysis(excelWrapper, this.ColumnTextBox.Text, this.FindValueInput.Text, this.CaseSensitiveCheck.IsChecked == true);
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

                var excelWrappers = filePaths.Select(filePath => new ExcelWrapper(filePath)).ToList();

                this.FindAnalysis(excelWrappers.ToArray(), this.ColumnTextBox.Text, this.FindValueInput.Text, this.CaseSensitiveCheck.IsChecked == true);
            }
        }

        private void FindAnalysis(ExcelWrapper excelWrapper, string column, string value, bool caseSensitive)
        {
            if (this.SpecificColumnCheckBox.IsChecked == false)
            {
                int columnCount = excelWrapper.GetColumnCount();

                for (int i = 0; i < columnCount; i++)
                {
                    string[] rows = excelWrapper.GetStringRows(i);

                    if (rows == null)
                    {
                        MessageBox.Show($"There isn't {column} column in {excelWrapper.FileName}");
                        return;
                    }

                    var output = new List<string>();

                    for (int j = 0; j < rows.Length; j++)
                    {
                        string rowValue = rows[j];

                        if (rowValue.Equals(value))
                        {
                            output.Add($"\"{value}\" was found at {ExcelWrapper.ConvertStringColumnToNumber(i)}{j}");
                        }
                        else if (!caseSensitive && rowValue.ToLower().Equals(value))
                        {
                            output.Add($"\"{value}\" was found at {ExcelWrapper.ConvertStringColumnToNumber(i)}{j}");
                        }
                    }

                    if (output.Count > 0)
                    {
                        this.Logger.Log(string.Join("\r\n", output));
                    }
                }
            }
            else
            {
                int columnNumber = ExcelWrapper.ConvertStringColumnToNumber(column);

                if (columnNumber == -1)
                {
                    MessageBox.Show($"Column '{column}' is not a valid column!");
                    return;
                }

                string[] rows = excelWrapper.GetStringRows(columnNumber);

                if (rows == null)
                {
                    MessageBox.Show($"There isn't {column} column in {excelWrapper.FileName}");
                    return;
                }

                var output = new List<string>();

                for (int j = 0; j < rows.Length; j++)
                {
                    string rowValue = rows[j];

                    if (rowValue.Equals(value))
                    {
                        output.Add($"\"{value}\" was found at {column}{j}");
                    }
                    else if (!caseSensitive && rowValue.ToLower().Equals(value))
                    {
                        output.Add($"\"{value}\" was found at {column}{j}");
                    }
                }

                if (output.Count > 0)
                {
                    this.Logger.Log(string.Join("\r\n", output));
                }
            }
        }

        private void FindAnalysis(ExcelWrapper[] excelWrappers, string column, string value, bool caseSensitive)
        {
            var logs = new List<string>();
            foreach (var excelWrapper in excelWrappers)
            {
                if (this.SpecificColumnCheckBox.IsChecked == false)
                {
                    int columnCount = excelWrapper.GetColumnCount();

                    for (int i = 0; i < columnCount; i++)
                    {
                        string[] rows = excelWrapper.GetStringRows(i);

                        if (rows == null)
                        {
                            MessageBox.Show($"There isn't {column} column in {excelWrapper.FileName}");
                            return;
                        }

                        var output = new List<string>();

                        for (int j = 0; j < rows.Length; j++)
                        {
                            string rowValue = rows[j];

                            if (rowValue.Equals(value))
                            {
                                output.Add($"\"{value}\" was found at {ExcelWrapper.ConvertStringColumnToNumber(i)}{j}");
                            }
                            else if (!caseSensitive && rowValue.ToLower().Equals(value))
                            {
                                output.Add($"\"{value}\" was found at {ExcelWrapper.ConvertStringColumnToNumber(i)}{j}");
                            }
                        }

                        if (output.Count > 0)
                        {
                            logs.AddRange(output);
                        }
                    }
                }
                else
                {
                    int columnNumber = ExcelWrapper.ConvertStringColumnToNumber(column);

                    if (columnNumber == -1)
                    {
                        MessageBox.Show($"Column '{column}' is not a valid column!");
                        return;
                    }

                    string[] rows = excelWrapper.GetStringRows(columnNumber);

                    if (rows == null)
                    {
                        MessageBox.Show($"There isn't {column} column in {excelWrapper.FileName}");
                        return;
                    }

                    var output = new List<string>();

                    for (int j = 0; j < rows.Length; j++)
                    {
                        string rowValue = rows[j];

                        if (rowValue.Equals(value))
                        {
                            output.Add($"\"{value}\" was found at {column}{j}");
                        }
                        else if (!caseSensitive && rowValue.ToLower().Equals(value))
                        {
                            output.Add($"\"{value}\" was found at {column}{j}");
                        }
                    }

                    if (output.Count > 0)
                    {
                        logs.AddRange(output);
                    }
                }
            }

            this.Logger.Log(string.Join("\r\n", logs));
        }

        public class BooleanToVisibilityConverter : IValueConverter
        {
            public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            {
                if (value is Visibility visibility)
                {
                    return visibility == Visibility.Visible;
                }

                return Visibility.Collapsed;
            }

            public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            {
                if (value is bool isVisible)
                {
                    return isVisible ? Visibility.Visible : Visibility.Collapsed;
                }

                return Visibility.Visible;
            }
        }
    }
}
