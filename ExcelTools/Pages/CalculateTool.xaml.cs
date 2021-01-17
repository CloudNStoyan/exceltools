using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ExcelTools.Attributes;

namespace ExcelTools.Pages
{
    [PageInfo(Header = "Calculate Tool")]
    public partial class CalculateTool : Page
    {
        public CalculateTool()
        {
            this.InitializeComponent();
        }

        private void RunMultipleFilesAnalysis(object sender, RoutedEventArgs e)
        {
            if (this.FileSelection.SelectedFiles == null)
            {
                AlertManager.NoFileSelected();
                return;
            }

            string[] filePaths = this.FileSelection.SelectedFiles;

            double sum = 0;

            foreach (string filePath in filePaths)
            {
                var excelWrapper = new ExcelWrapper(filePath);

                string firstColumn = this.FirstColumn.Text;

                int firstColumnNumber = ExcelWrapper.ConvertStringColumnToNumber(firstColumn);

                double[] firstColumnValues = excelWrapper.GetDoubleRows(firstColumnNumber);

                string fileName = Path.GetFileName(filePath);

                if (firstColumnValues == null)
                {
                    MessageBox.Show($"There is no column '{firstColumn}' in {fileName}");
                    return;
                }

                sum += firstColumnValues.Sum();
            }

            this.MultipleFilesOutput.Text = $"Summation of all excel files\n{sum}";
        }

        private void RunColumnAnalysis(object sender, RoutedEventArgs e)
        {
            if (this.FileSelection.SelectedFiles == null)
            {
                AlertManager.NoFileSelected();
                return;
            }

            var operation = (OperationType) this.OperationSelector.SelectedIndex;

            string[] filePaths = this.FileSelection.SelectedFiles;

            var output = new List<string>();

            foreach (string filePath in filePaths)
            {
                var excelWrapper = new ExcelWrapper(filePath);

                string firstColumn = this.FirstColumn.Text;
                string secondColumn = this.SecondColumn.Text;

                int firstColumnNumber = ExcelWrapper.ConvertStringColumnToNumber(firstColumn);
                int secondColumnNumber = ExcelWrapper.ConvertStringColumnToNumber(secondColumn);

                double[] firstColumnValues = excelWrapper.GetDoubleRows(firstColumnNumber);
                double[] secondColumnValues = excelWrapper.GetDoubleRows(secondColumnNumber);

                string fileName = Path.GetFileName(filePath);

                if (firstColumnValues == null)
                {
                    MessageBox.Show($"There is no column '{firstColumn}' in {fileName}");
                    return;
                }

                if (secondColumnValues == null)
                {
                    MessageBox.Show($"There is no column '{secondColumn}' in {fileName}");
                    return;
                }

                double[] newValues = new double[firstColumnValues.Length > secondColumnValues.Length
                    ? firstColumnValues.Length
                    : secondColumnValues.Length];

                switch (operation)
                {
                    case OperationType.Divide:
                        for (int i = 0; i < newValues.Length; i++)
                        {
                            newValues[i] = firstColumnValues[i] / secondColumnValues[i];
                        }
                        break;
                    case OperationType.Multiply:
                        for (int i = 0; i < newValues.Length; i++)
                        {
                            newValues[i] = firstColumnValues[i] * secondColumnValues[i];
                        }
                        break;
                    case OperationType.Add:
                        for (int i = 0; i < newValues.Length; i++)
                        {
                            newValues[i] = firstColumnValues[i] + secondColumnValues[i];
                        }
                        break;
                    case OperationType.Subtract:
                        for (int i = 0; i < newValues.Length; i++)
                        {
                            newValues[i] = firstColumnValues[i] - secondColumnValues[i];
                        }
                        break;
                }

                output.Add($"Results from '{fileName}':");
                output.Add(string.Join("\n", newValues));
            }

            this.ColumnOutput.Text = string.Join("\n", output);
        }

        private enum OperationType
        {
            Divide,
            Multiply,
            Add,
            Subtract
        }
    }
}
