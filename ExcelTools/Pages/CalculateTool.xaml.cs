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

        private void FileSelection_OnFileSelected()
        {

        }

        private void RunAnalysis(object sender, RoutedEventArgs e)
        {
            var operation = (OperationType) this.OperationSelector.SelectedIndex;

            var excelWrapper = new ExcelWrapper(this.FileSelection.SelectedFiles[0]);

            int firstColumn = ExcelWrapper.ConvertStringColumnToNumber(this.FirstColumn.Text);
            int secondColumn = ExcelWrapper.ConvertStringColumnToNumber(this.SecondColumn.Text);

            double[] firstColumnValues = excelWrapper.GetDoubleRows(firstColumn);
            double[] secondColumnValues = excelWrapper.GetDoubleRows(secondColumn);

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

            this.Output.Text = string.Join("\n", newValues);
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
