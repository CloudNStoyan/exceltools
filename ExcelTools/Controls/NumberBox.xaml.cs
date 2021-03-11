using System;
using System.Windows;
using System.Windows.Controls;

namespace ExcelTools.Controls
{
    public partial class NumberBox : UserControl
    {
        public int Min
        {
            get => (int)this.GetValue(MinProperty);
            set => this.SetValue(MinProperty, value);
        }

        public static readonly DependencyProperty MinProperty
            = DependencyProperty.Register(
                nameof(Min),
                typeof(int),
                typeof(NumberBox),
                new PropertyMetadata(0)
            );

        public int Max
        {
            get => (int)this.GetValue(MaxProperty);
            set => this.SetValue(MaxProperty, value);
        }

        public static readonly DependencyProperty MaxProperty
            = DependencyProperty.Register(
                nameof(Max),
                typeof(int),
                typeof(NumberBox),
                new PropertyMetadata(1)
            );

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
                this.Input.Text = this.Min.ToString();
                return;
            }

            if (!int.TryParse(text, out int result))
            {
                this.Input.Text = this.Text ?? this.Min.ToString();
                this.Input.CaretIndex = this.Input.Text.Length;
                return;
            }

            if (result > this.Max)
            {
                this.Input.Text = this.Max.ToString();
                this.Input.CaretIndex = this.Input.Text.Length;
                return;
            }

            this.Text = text;
            this.Number = result;

            this.Input.Text = this.Input.Text.Trim();

            if (this.Input.Text != this.Min.ToString())
            {
                this.Input.Text = this.Input.Text.TrimStart('0');
            }

            this.Input.CaretIndex = this.Input.Text.Length;

            this.NumberChanged?.Invoke();
        }
    }
}
