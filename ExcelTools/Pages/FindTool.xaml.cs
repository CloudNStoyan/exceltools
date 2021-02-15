using System;
using System.Collections.Generic;
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
            string[] columns = this.SpecificColumnCheckBox.IsChecked == true ? this.ColumnTextBox.GetColumns() : null;

            bool multipleFiles = this.FileSelection.MultipleFilesChecked;

            string[] filePaths =
                multipleFiles ? this.FileSelection.SelectedFiles : new[] {this.FileSelection.SelectedFile};

            this.FindInExcel(value, caseSensitive, columns, filePaths);
        }

        private void FindInExcel(string value, bool caseSensitive, string[] columns, string[] filePaths)
        {
            if (filePaths == null || filePaths.Length == 0)
            {
                AlertManager.NoFileSelected();
                return;
            }

            var excelWrappers = filePaths.Select(filePath => new ExcelWrapper(filePath)).ToArray();

            this.FindAnalysis(excelWrappers, value, caseSensitive, columns);
        }

        private string[] CheckColumns(string[] columns, ExcelWrapper excelWrapper, string value, bool caseSensitive)
        {
            var output = new List<string>();

            foreach (string columnNumber in columns)
            {
                int column = ExcelWrapper.ConvertStringColumnToNumber(columnNumber);

                string[] rows = excelWrapper.GetStringRows(column);

                if (rows != null)
                {
                    for (int j = 0; j < rows.Length; j++)
                    {
                        string rowValue = rows[j];

                        if (rowValue != null && rowValue.Equals(value))
                        {
                            output.Add($"\"{rowValue}\" was found at {columnNumber}{j}");
                        }
                        else if (rowValue != null && !caseSensitive && string.Equals(rowValue, value, StringComparison.CurrentCultureIgnoreCase))
                        {
                            output.Add($"\"{rowValue}\" was found at {columnNumber}{j}");
                        }
                    }
                }
                else
                {

                    AlertManager.Custom($"There isn't {columnNumber} column in {excelWrapper.FileName}");
                }
            }

            return output.ToArray();
        }

        private const string NoMatchesFound = "No matches were found!";
        private void FindAnalysis(ExcelWrapper[] excelWrappers, string value, bool caseSensitive, string[] columns)
        {
            var logs = new List<string>();

            foreach (var excelWrapper in excelWrappers)
            {
                if (columns == null)
                {
                    columns = excelWrapper.GetColumns();
                }

                logs.AddRange(this.CheckColumns(columns, excelWrapper, value, caseSensitive));
            }

            this.Logger.Log(logs.Count > 0 ? string.Join("\r\n", logs) : NoMatchesFound);
        }
    }
}
