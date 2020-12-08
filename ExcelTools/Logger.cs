using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace ExcelTools
{
    public class Logger
    {
        private StackPanel StackPanel { get; }
        private bool UseTimestamp { get; }

        public Logger(StackPanel stackPanel, bool useTimestamp)
        {
            this.StackPanel = stackPanel;
            this.UseTimestamp = useTimestamp;
        }

        public void Log(string text)
        {
            var logWrapper = new StackPanel();

            var textBlock = new TextBlock {Text = text, TextWrapping = TextWrapping.Wrap};

            if (this.UseTimestamp)
            {
                var date = DateTime.Now;

                string timestamp = $"[{date.Hour.ToString().PadLeft(2, '0')}:{date.Minute.ToString().PadLeft(2, '0')}:{date.Second.ToString().PadLeft(2,'0')}]";

                var timeStampTextBlock = new TextBlock
                {
                    Text = timestamp,
                    Foreground = Brushes.DimGray
                };

                logWrapper.Children.Add(timeStampTextBlock);
            }

            logWrapper.Children.Add(textBlock);

            this.StackPanel.Children.Add(logWrapper);
        }
    }
}