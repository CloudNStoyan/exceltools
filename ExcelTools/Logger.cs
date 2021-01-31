using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace ExcelTools
{
    public class Logger
    {
        private StackPanel StackPanel { get; }
        private bool UseTimestamp { get; }
        public string LogText => string.Join("\r\n", this.logData);
        private readonly List<string> logData = new List<string>();

        public Logger(StackPanel stackPanel, bool useTimestamp)
        {
            this.StackPanel = stackPanel;
            this.UseTimestamp = useTimestamp;
        }

        public void Clear()
        {
            this.StackPanel.Children.Clear();
            this.logData.Clear();
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

                this.logData.Add($"{timestamp}\r\n{text}");
            }

            logWrapper.Children.Add(textBlock);

            this.StackPanel.Children.Add(logWrapper);
        }
    }
}