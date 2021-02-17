using System;
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
            var operation = (Operation) ((ComboBoxItem) this.OperationDropdown.SelectedItem).DataContext;

            switch (operation)
            {
                case Operation.Concat:
                    this.ConcatOperation();
                    break;
            }
        }

        private void ConcatOperation()
        {
            this.Output.OutputTextBox.Text = "test";
        }
    }
}