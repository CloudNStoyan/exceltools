using System.Windows;
using System.Windows.Controls;

namespace ExcelTools.Controls
{
    public partial class Input : UserControl
    {
        public Input() => this.InitializeComponent();

        private void MainTextBox_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            string text = this.MainTextBox.Text;

            this.MainTextBox.Text = text.Length > 1 ? text[text.Length - 1].ToString().ToUpper() : text.ToUpper();

            this.MainTextBox.CaretIndex = 1;
        }
    }
}
