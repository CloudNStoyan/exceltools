using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
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

                string[] column = this.Columns.GetColumns();

                var data = this.Export(excelWrapper, column, this.SkipEmpty.IsChecked == true);
                
                if (data == null)
                {
                    MessageBox.Show($"There is no column {column} in {excelWrapper.FileName}");
                    return;
                }

                data = data.Select(x => x.Select(y => y.Replace("\n", " ")).ToArray()).ToArray();

                int separatorIndex = int.Parse(this.SeparatorsPanel.Children.OfType<RadioButton>()
                    .First(r => r.IsChecked == true).DataContext.ToString());

                string separator = this.separators[separatorIndex];

                this.OutputTextbox.Text = string.Join(separator, data.Select(x => string.Join("\t",x)));
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

                string[] column = this.Columns.GetColumns();

                var data = this.Export(excelWrappers, column, this.SkipEmpty.IsChecked == true);
                
                if (data == null)
                {
                    MessageBox.Show($"No data can be exported");
                    return;
                }

                data = data.Select(x => x.Select(y => y.Select(z => z.Replace("\n", " ")).ToArray()).ToArray()).ToArray();

                int separatorIndex = int.Parse(this.SeparatorsPanel.Children.OfType<RadioButton>()
                    .First(r => r.IsChecked == true).DataContext.ToString());

                string separator = this.separators[separatorIndex];

                this.OutputTextbox.Text = string.Join(separator, data.Select(x => x));
            }
        }

        private string[][] Export(ExcelWrapper excelWrapper, string[] column, bool skipEmpty = false)
        {
            int[] columnNumbers = column.Select(ExcelWrapper.ConvertStringColumnToNumber).ToArray();

            for (var i = 0; i < columnNumbers.Length; i++)
            {
                int columnNumber = columnNumbers[i];
                if (columnNumber == -1)
                {
                    MessageBox.Show($"Column '{column[i]}' is not a valid column, and it will be skipped!");
                }
            }

            columnNumbers = columnNumbers.Where(x => x != -1).ToArray();

            if (columnNumbers.Length < 1)
            {
                MessageBox.Show("There are no valid columns!");
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
                    MessageBox.Show($"Column '{column[i]}' is not a valid column, and it will be skipped!");
                }
            }

            columnNumbers = columnNumbers.Where(x => x != -1).ToArray();

            if (columnNumbers.Length < 1)
            {
                MessageBox.Show("There are no valid columns!");
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