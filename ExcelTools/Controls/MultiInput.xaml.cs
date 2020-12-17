using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace ExcelTools.Controls
{
    public partial class MultiInput : UserControl
    {
        public MultiInput() => this.InitializeComponent();
        private const int MaxInputs = 10;
        private int currentInputs = 1;

        public string[] GetColumns() => this.InputPanel.Children.OfType<Input>().Select(x => x.Text).ToArray();

        private void AddInput(object sender, RoutedEventArgs e)
        {
            this.InputPanel.Children.Add(new Input {Margin = new Thickness(5,0,0,0), Text = "A", Width = 25, Height = 25});
            this.currentInputs++;

            if (this.currentInputs == MaxInputs)
            {
                this.AddBtn.IsEnabled = false;
            }

            if (this.currentInputs > 1)
            {
                this.RemoveBtn.IsEnabled = true;
            }

        }

        private void RemoveInput(object sender, RoutedEventArgs e)
        {
            this.InputPanel.Children.RemoveAt(this.InputPanel.Children.Count - 1);
            this.currentInputs--;

            if (this.currentInputs == 1)
            {
                this.RemoveBtn.IsEnabled = false;
            }

            if (this.currentInputs < MaxInputs)
            {
                this.AddBtn.IsEnabled = true;
            }
        }
    }
}
