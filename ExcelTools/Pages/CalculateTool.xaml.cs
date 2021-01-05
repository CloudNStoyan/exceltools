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

            switch (operation)
            {
                case OperationType.Divide:
                    var excelWrapper = new ExcelWrapper(this.FileSelection.SelectedFiles[0]);

                    int firstColumn = ExcelWrapper.ConvertStringColumnToNumber(this.FirstColumn.Text);
                    int secondColumn = ExcelWrapper.ConvertStringColumnToNumber(this.SecondColumn.Text);

                    double[] firstColumnValues = excelWrapper.GetDoubleRows(firstColumn);
                    double[] secondColumnValues = excelWrapper.GetDoubleRows(secondColumn);

                    double[] newValues = new double[firstColumnValues.Length > secondColumnValues.Length
                        ? firstColumnValues.Length
                        : secondColumnValues.Length];

                    for (int i = 0; i < newValues.Length; i++)
                    {
                        newValues[i] = firstColumnValues[i] / secondColumnValues[i];
                    }

                    this.Output.Text = string.Join("\n", newValues);

                    break;
                case OperationType.Multiply:
                    break;
                case OperationType.Add:
                    break;
                case OperationType.Subtract:
                    break;
            }
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
