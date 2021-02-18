using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using ExcelTools.Attributes;

namespace ExcelTools.Pages
{
    [PageInfo(Header = "Multifiles Operation")]
    public partial class MultifilesOperation : Page
    {
        private enum Operation
        {
            Concat,
            Sum
        }

        public MultifilesOperation()
        {
            this.InitializeComponent();

            this.SetupUI();
        }

        private void SetupUI()
        {
            string[] operations = Enum.GetNames(typeof(Operation));

            foreach (string operation in operations)
            {
                Enum.TryParse(operation, out Operation operationEnum);

                this.OperationDropdown.Items.Add(new ComboBoxItem { Content = new TextBlock { Text = operation }, DataContext = operationEnum});
            }

            this.OperationDropdown.SelectedIndex = 0;
        }

        private void RunOperation(object sender, RoutedEventArgs e)
        {
            this.Output.OutputTextBox.Clear();

            var operation = (Operation) ((ComboBoxItem) this.OperationDropdown.SelectedItem).DataContext;

            switch (operation)
            {
                case Operation.Concat:
                    this.ConcatOperation();
                    break;
                case Operation.Sum:
                    this.SumOperation();
                    break;
            }
        }

        private void SumOperation()
        {
            var excelWrappers = this.FileSelection.SelectedFiles.Select(path => new ExcelWrapper(path)).ToArray();

            string column = this.ColumnInput.Text;
            int columnNumber = ExcelWrapper.ConvertStringColumnToNumber(column);

            double sum = excelWrappers.Sum(x => x.GetDoubleRows(columnNumber).Sum());

            this.Output.OutputTextBox.Text = sum.ToString("F2");
        }

        private void ConcatOperation()
        {
            var excelWrappers = this.FileSelection.SelectedFiles.Select(path => new ExcelWrapper(path));

            string column = this.ColumnInput.Text;
            int columnNumber = ExcelWrapper.ConvertStringColumnToNumber(column);

            var stringBuilder = new StringBuilder();

            foreach (var excelWrapper in excelWrappers)
            {
                string[] entries = excelWrapper.GetStringRows(columnNumber);

                foreach (string entry in entries)
                {
                    stringBuilder.Append(entry);
                }
            }

            this.Output.OutputTextBox.Text = stringBuilder.ToString();
        }
    }
}