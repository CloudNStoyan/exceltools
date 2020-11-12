using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;

namespace ExcelTools
{
    public partial class MainWindow : Window
    {
        private Logger Logger { get; }
        public MainWindow()
        {
            this.InitializeComponent();

            this.Logger = new Logger(this.LogStackPanel, true);

            this.SetupPages();
        }

        private void SetupPages()
        {
            var pageClasses = Assembly.GetExecutingAssembly().DefinedTypes.Where(x => x.Namespace == "ExcelTools.Pages" && x.BaseType?.Name == "Page").ToArray();

            var pagePairs = (from typeInfo in pageClasses
                let instance = (Page)Activator.CreateInstance(typeInfo, this.Logger)
                let header = (string)typeInfo.GetDeclaredField("Header").GetValue(instance)
                select new KeyValuePair<string, Page>(header, instance)).ToArray();

            foreach (var pair in pagePairs)
            {
                var button = new Button
                {
                    Content = pair.Key,
                    Padding = new Thickness(5),
                    Margin = new Thickness(5)
                };

                button.Click += (sender, args) => { this.Settings.Navigate(pair.Value); };

                this.NavigationContainer.Children.Add(button);
            }
        }

        private void ClearLogStackPanelHandler(object sender, RoutedEventArgs e)
        {
            this.LogStackPanel.Children.Clear();
        }
    }
}