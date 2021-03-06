﻿using System;
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

        public ExportTool() => this.InitializeComponent();

        private readonly string[] separators = {" ", Environment.NewLine, ","};

        private void RunAnalysis(object sender, RoutedEventArgs e)
        {
            int separatorIndex = int.Parse(this.SeparatorsPanel.Children.OfType<RadioButton>()
                .First(r => r.IsChecked == true).DataContext.ToString());

            string separator = this.separators[separatorIndex];

            if (this.FileSelection.MultipleFilesChecked == false)
            {
                string filePath = this.FileSelection.SelectedFile;

                this.ExportAnalysis(new[] {filePath}, separator);
            }
            else
            {
                string[] filePaths = this.FileSelection.SelectedFiles;

                this.ExportAnalysis(filePaths, separator);
            }
        }

        private void ExportAnalysis(string[] filePaths, string separator)
        {
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

        private const string NoValidColumns = "There are no valid columns!";

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
                AlertManager.Custom(NoValidColumns);
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