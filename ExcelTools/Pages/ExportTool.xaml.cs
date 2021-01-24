using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ExcelTools.Attributes;

namespace ExcelTools.Pages
{
    [PageInfo(Header = "Export Tool", Order = 0)]
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
                    AlertManager.NoFileSelected();
                    return;
                }

                var excelWrapper = new ExcelWrapper(filePath);

                string[] columns = this.SpecificColumnCheckBox.IsChecked == true ? this.Columns.GetColumns() : excelWrapper.GetColumns();

                var data = this.Export(excelWrapper, columns, this.SkipEmpty.IsChecked == true);
                
                if (data == null)
                {
                    AlertManager.Custom($"There is no column {columns} in {excelWrapper.FileName}");
                    return;
                }

                data = data.Select(x => x.Select(y => y.Replace("\n", " ")).ToArray()).ToArray();

                int separatorIndex = int.Parse(this.SeparatorsPanel.Children.OfType<RadioButton>()
                    .First(r => r.IsChecked == true).DataContext.ToString());

                string separator = this.separators[separatorIndex];
                

                this.Output.OutputTextBox.Text = string.Join(separator, this.ConvertDataToTable(data));
            }
            else
            {
                string[] filePaths = this.FileSelection.SelectedFiles;

                if (filePaths == null || filePaths.Length == 0)
                {
                    AlertManager.NoFileSelected();
                    return;
                }

                foreach (string selectedFile in filePaths)
                {
                    if (!File.Exists(selectedFile))
                    {
                        AlertManager.Custom($"Cannot find {selectedFile}!");
                        return;
                    }
                }

                var excelWrappers = filePaths.Select(filePath => new ExcelWrapper(filePath)).ToArray();

                string[] column = this.Columns.GetColumns();

                var data = this.Export(excelWrappers, column, this.SkipEmpty.IsChecked == true);
                
                if (data == null)
                {
                    AlertManager.Custom("No data can be exported");
                    return;
                }

                data = data.Select(x => x.Select(y => y.Select(z => z.Replace("\n", " ")).ToArray()).ToArray()).ToArray();

                int separatorIndex = int.Parse(this.SeparatorsPanel.Children.OfType<RadioButton>()
                    .First(r => r.IsChecked == true).DataContext.ToString());

                string separator = this.separators[separatorIndex];

                var output = new List<string>();

                for (int i = 0; i < data.Length; i++)
                {
                    if (this.AddFileName.IsChecked == true)
                    {
                        output.Add("====");
                        output.Add(Path.GetFileName(filePaths[i]));
                    }
                    
                    output.Add(string.Join(separator, this.ConvertDataToTable(data[i])));
                }

                this.Output.OutputTextBox.Text = string.Join("\r\n", output);
            }
        }

        private string[] ConvertDataToTable(string[][] data)
        {
            int maxLength = 0;

            foreach (var column in data)
            {
                if (column.Length > maxLength)
                {
                    maxLength = column.Length;
                }
            }

            string[,] dataMatrix = new string[maxLength, data.Length];

            for (int i = 0; i < data.Length; i++)
            {
                string[] row = data[i];

                for (int j = 0; j < row.Length; j++)
                {
                    dataMatrix[j, i] = row[j];
                }
            }

            var output = new List<string>();

            for (int i = 0; i < dataMatrix.GetLength(0); i++)
            {
                string line = "";
                for (int j = 0; j < dataMatrix.GetLength(1); j++)
                {
                    if (j != 0)
                    {
                        line += "\t" + dataMatrix[i, j];
                    }
                    else
                    {
                        line = dataMatrix[i, j];
                    }
                }
                output.Add(line);
            }

            return output.ToArray();
        }

        private string[][] Export(ExcelWrapper excelWrapper, string[] column, bool skipEmpty = false)
        {
            int[] columnNumbers = column.Select(ExcelWrapper.ConvertStringColumnToNumber).ToArray();

            for (var i = 0; i < columnNumbers.Length; i++)
            {
                int columnNumber = columnNumbers[i];
                if (columnNumber == -1)
                {
                    AlertManager.Custom($"Column '{column[i]}' is not a valid column, and it will be skipped!");
                }
            }

            columnNumbers = columnNumbers.Where(x => x != -1).ToArray();

            if (columnNumbers.Length < 1)
            {
                AlertManager.Custom("There are no valid columns!");
                return null;
            }


            string[][] columnData = !skipEmpty
                ? columnNumbers.Select(excelWrapper.GetStringRows).ToArray()
                : columnNumbers.Select(excelWrapper.GetStringRows)
                    .Select(data => data.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray()).ToArray();

            return columnData;
        }

        private string[][][] Export(ExcelWrapper[] excelWrappers, string[] column, bool skipEmpty = false)
        {
            int[] columnNumbers = column.Select(ExcelWrapper.ConvertStringColumnToNumber).ToArray();

            for (var i = 0; i < columnNumbers.Length; i++)
            {
                int columnNumber = columnNumbers[i];
                if (columnNumber == -1)
                {
                    AlertManager.Custom($"Column '{column[i]}' is not a valid column, and it will be skipped!");
                }
            }

            columnNumbers = columnNumbers.Where(x => x != -1).ToArray();

            if (columnNumbers.Length < 1)
            {
                AlertManager.Custom("There are no valid columns!");
                return null;
            }

            return (from excelWrapper in excelWrappers
                select !skipEmpty
                    ? columnNumbers.Select(excelWrapper.GetStringRows).ToArray()
                    : columnNumbers.Select(excelWrapper.GetStringRows)
                        .Select(data => data.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray())
                        .ToArray()).ToArray();
        }
    }
}