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
        public string LogText => string.Join("\r\n", this.LogData);
        private List<string> LogData = new List<string>();

        public Logger(StackPanel stackPanel, bool useTimestamp)
        {
            this.StackPanel = stackPanel;
            this.UseTimestamp = useTimestamp;
        }

        public void Clear()
        {
            this.StackPanel.Children.Clear();
            this.LogData.Clear();
        }

        public void Log(string[] logs)
        {
            foreach (string log in logs)
            {
                this.Log(log);
            }
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

                this.LogData.Add($"{timestamp}\r\n{text}");
            }

            logWrapper.Children.Add(textBlock);

            this.StackPanel.Children.Add(logWrapper);
        }
    }
}