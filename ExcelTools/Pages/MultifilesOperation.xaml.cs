using System;
using System.Windows.Controls;
using ExcelTools.Attributes;

namespace ExcelTools.Pages
{
    [PageInfo(Header = "Multifiles Operation")]
    public partial class MultifilesOperation : Page
    {
        private enum Operation
        {
            Concat
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
                this.OperationDropdown.Items.Add(new ComboBoxItem { Content = new TextBlock { Text = operation } });
            }

            this.OperationDropdown.SelectedIndex = 0;
        }
    }
}