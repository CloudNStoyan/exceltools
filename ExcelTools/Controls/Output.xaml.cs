using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace ExcelTools.Controls
{
    public partial class Output : UserControl
    {
        public Output()
        {
            this.InitializeComponent();
        }

        private void SaveOutput(object sender, RoutedEventArgs e)
        {
            string text = this.OutputTextBox.Text;

            var saveFileDialog = new SaveFileDialog
            {
                DefaultExt = ".txt", InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                FileName = "export-data.txt"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                File.WriteAllText(saveFileDialog.FileName, text);
            }
        }
    }
}
