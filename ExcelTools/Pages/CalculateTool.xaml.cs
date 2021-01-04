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
