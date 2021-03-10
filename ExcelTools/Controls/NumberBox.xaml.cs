using System;
using System.Windows;
using System.Windows.Controls;

namespace ExcelTools.Controls
{
    public partial class NumberBox : UserControl
    {
        public NumberBox()
        {
            this.InitializeComponent();
        }

        private void Increment(object sender, RoutedEventArgs e) => this.Input.Text = (this.Number + 1).ToString();

        private void Decrement(object sender, RoutedEventArgs e) => this.Input.Text = (this.Number - 1).ToString();

        private string Text { get; set; }
        public int Number = 1;

        public event Action NumberChanged;

        private void Input_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            string text = this.Input.Text;

            if (string.IsNullOrWhiteSpace(text))
            {
                this.Input.Text = "1";
                return;
            }

            if (!int.TryParse(text, out int result))
            {
                this.Input.Text = this.Text ?? "1";
                this.Input.CaretIndex = this.Input.Text.Length;
                return;
            }

            this.Text = text;
            this.Number = result;

            this.Input.Text = this.Input.Text.Trim();

            if (this.Input.Text != "1")
            {
                this.Input.Text = this.Input.Text.TrimStart('0');
            }

            this.Input.CaretIndex = this.Input.Text.Length;

            this.NumberChanged?.Invoke();
        }
    }
}
