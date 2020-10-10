using System;
using System.Windows;
using System.Windows.Controls;

namespace ExcelTools
{
    public class Logger
    {
        private StackPanel StackPanel { get; }
        private bool UseTimestamp { get; }

        public Logger(StackPanel stackPanel)
        {
            this.StackPanel = stackPanel;
        }

        public Logger(StackPanel stackPanel, bool useTimestamp)
        {
            this.StackPanel = stackPanel;
            this.UseTimestamp = useTimestamp;
        }

        public void Log(string text)
        {
            var textBlock = new TextBlock {Text = text, TextWrapping = TextWrapping.Wrap};

            if (this.UseTimestamp)
            {
                var date = DateTime.Now;

                string timestamp = $"[{date.Hour}:{date.Minute}:{date.Second}]";

                textBlock.Text = $"{timestamp} {text}";
            }

            this.StackPanel.Children.Add(textBlock);
        }
    }
}