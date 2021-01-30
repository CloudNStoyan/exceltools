using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
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
                AlertManager.Custom("Value cannot be empty or white space only!");
                return;
            }

            string value = this.FindValueInput.Text;
            bool caseSensitive = this.CaseSensitiveCheck.IsChecked == true;
            string[] columns = this.ColumnTextBox.GetColumns();

            if (this.FileSelection.MultipleFilesChecked == false)
            {
                string filePath = this.FileSelection.SelectedFile;

                if (!File.Exists(filePath))
                {
                    AlertManager.NoFileSelected();
                    return;
                }

                var excelWrapper = new ExcelWrapper(filePath);

                if (this.SpecificColumnCheckBox.IsChecked == true)
                {
                    this.FindAnalysis(excelWrapper, value, caseSensitive, columns);
                }
                else
                {
                    this.FindAnalysis(excelWrapper, value, caseSensitive);
                }
            }
            else
            {
                string[] filePaths = this.FileSelection.SelectedFiles;

                if (filePaths == null || filePaths.Length == 0)
                {
                    AlertManager.NoFileSelected();
                    return;
                }

                var excelWrappers = filePaths.Select(filePath => new ExcelWrapper(filePath)).ToArray();

                if (this.SpecificColumnCheckBox.IsChecked == true)
                {
                    this.FindAnalysis(excelWrappers, value, caseSensitive, columns);
                }
                else
                {
                    this.FindAnalysis(excelWrappers, value, caseSensitive);
                }
            }
        }

        private string[] CheckColumns(string[] columns, ExcelWrapper excelWrapper, string value, bool caseSensitive)
        {
            var output = new List<string>();

            for (int i = 0; i < columns.Length; i++)
            {
                int column = ExcelWrapper.ConvertStringColumnToNumber(columns[i]);

                string[] rows = excelWrapper.GetStringRows(column);

                if (rows != null)
                {
                    for (int j = 0; j < rows.Length; j++)
                    {
                        string rowValue = rows[j];

                        if (rowValue != null && rowValue.Equals(value))
                        {
                            output.Add($"\"{rowValue}\" was found at {columns[i]}{j}");
                        }
                        else if (rowValue != null && !caseSensitive && string.Equals(rowValue, value, StringComparison.CurrentCultureIgnoreCase))
                        {
                            output.Add($"\"{rowValue}\" was found at {columns[i]}{j}");
                        }
                    }
                }
                else
                {

                    AlertManager.Custom($"There isn't {columns[i]} column in {excelWrapper.FileName}");
                }
            }

            return output.ToArray();
        }

        private void FindAnalysis(ExcelWrapper excelWrapper, string value, bool caseSensitive)
        {
            string[] columns = excelWrapper.GetColumns();

            string[] logs = this.CheckColumns(columns, excelWrapper, value, caseSensitive);

            this.Logger.Log(logs.Length > 0 ? string.Join("\r\n", logs) : CustomResources.NoMatchesFound);
        }

        private void FindAnalysis(ExcelWrapper[] excelWrappers, string value, bool caseSensitive)
        {
            var logs = new List<string>();

            foreach (var excelWrapper in excelWrappers)
            {
                string[] columns = excelWrapper.GetColumns();

                logs.AddRange(this.CheckColumns(columns, excelWrapper, value, caseSensitive));
            }

            this.Logger.Log(logs.Count > 0 ? string.Join("\r\n", logs) : CustomResources.NoMatchesFound);
        }

        private void FindAnalysis(ExcelWrapper excelWrapper, string value, bool caseSensitive, string[] columns)
        {
            string[] logs = this.CheckColumns(columns, excelWrapper, value, caseSensitive);

            this.Logger.Log(logs.Length > 0 ? string.Join("\r\n", logs) : "No matches were found!");
        }

        private void FindAnalysis(ExcelWrapper[] excelWrappers, string value, bool caseSensitive, string[] columns)
        {
            var logs = new List<string>();

            foreach (var excelWrapper in excelWrappers)
            {
                logs.AddRange(this.CheckColumns(columns, excelWrapper, value, caseSensitive));
            }

            this.Logger.Log(logs.Count > 0 ? string.Join("\r\n", logs) : "No matches were found!");
        }
    }
}
