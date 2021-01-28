using System.Windows;
using System.Windows.Controls;

namespace ExcelTools.Controls
{
    public partial class Input : UserControl
    {
        public string Text
        {
            get => (string)this.GetValue(TextProperty);
            set => this.SetValue(TextProperty, value);
        }

        public static readonly DependencyProperty TextProperty
            = DependencyProperty.Register(
                nameof(Text),
                typeof(string),
                typeof(Input),
                new PropertyMetadata(string.Empty)
            );

        public Input() => this.InitializeComponent();

        private void MainTextBox_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            string text = this.MainTextBox.Text;

            if (string.IsNullOrWhiteSpace(text))
            {
                return;
            }

            char lastChar = text[text.Length - 1];

            this.MainTextBox.Text = char.IsLetter(lastChar) ? lastChar.ToString().ToUpper() : text.Replace(lastChar.ToString(), "").ToUpper();

            this.MainTextBox.CaretIndex = 1;
        }

        private void Input_OnLoaded(object sender, RoutedEventArgs e) => this.DataContext = this;
    }
}
