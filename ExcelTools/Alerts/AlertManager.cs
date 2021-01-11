using System;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace ExcelTools.Alerts
{
    public static class AlertManager
    {
        private static TextBlock AlertBody { get; set; }
        private static Grid AlertWrapper { get; set; }
        private static Timer Timer { get; set; }
        public static void SetupAlert(Grid alertWrapper, TextBlock alertBody)
        {
            AlertBody = alertBody;
            AlertWrapper = alertWrapper;

            Timer = new Timer
            {
                AutoReset = false,
                Interval = 3000
            };

            Timer.Elapsed += (sender, args) =>
            {
                Application.Current.Dispatcher.BeginInvoke(
                    DispatcherPriority.Background,
                    new Action(() =>
                    {
                        AlertWrapper.Visibility = Visibility.Hidden;
                    }));
            };
        }

        public static void NoFileSelected() => Create("You need to select file first");

        private static void Create(string text)
        {
            Application.Current.Dispatcher.BeginInvoke(
                DispatcherPriority.Background,
                new Action(() =>
                {
                    AlertBody.Text = text;
                    AlertWrapper.Visibility = Visibility.Visible;

                    Timer.Start();
                }));

        }
    }
}
