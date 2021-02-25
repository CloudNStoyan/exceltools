using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace ExcelTools.Controls
{
    public partial class Output : UserControl
    {
        public string DefaultExt
        {
            get => (string)this.GetValue(DefaultExtProperty);
            set => this.SetValue(DefaultExtProperty, value);
        }

        public static readonly DependencyProperty DefaultExtProperty
            = DependencyProperty.Register(
                nameof(DefaultExt),
                typeof(string),
                typeof(Output),
                new PropertyMetadata("json")
            );

        public double TextboxHeight
        {
            get => (double)this.GetValue(TextboxHeightProperty);
            set => this.SetValue(TextboxHeightProperty, value);
        }

        public static readonly DependencyProperty TextboxHeightProperty
            = DependencyProperty.Register(
                nameof(TextboxHeight),
                typeof(double),
                typeof(Output),
                new PropertyMetadata(60d)
            );

        public string FileName { get; set; }

        public Output() => this.InitializeComponent();

        private void SaveOutput(object sender, RoutedEventArgs e)
        {
            string text = this.OutputTextBox.Text;

            var saveFileDialog = new SaveFileDialog
            {
                DefaultExt = this.DefaultExt, InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                FileName = this.FileName,
                Filter = "JSON|*.json|TXT files (*.txt)|*.txt"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                File.WriteAllText(saveFileDialog.FileName, text);
            }
        }

        private void CopyToClipboard(object sender, RoutedEventArgs e)
        {
            string text = this.OutputTextBox.Text;

            Clipboard.SetText(text);

            AlertManager.Custom("JSON coppied to clipboard!");
        }

        private void ClearOutput(object sender, RoutedEventArgs e) => this.OutputTextBox.Clear();
    }
}
