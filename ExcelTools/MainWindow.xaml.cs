using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using ExcelTools.Attributes;

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

            foreach (var pageClass in pageClasses)
            {
                var instance = (Page) Activator.CreateInstance(pageClass, this.Logger);
                string header = pageClass.GetCustomAttribute<PageInfo>().Header;

                var button = new Button
                {
                    Content = header,
                    Padding = new Thickness(5),
                    Margin = new Thickness(5)
                };

                button.Click += (sender, args) => { this.Settings.Navigate(instance); };

                this.NavigationContainer.Children.Add(button);
            }
        }

        private void ClearLogStackPanelHandler(object sender, RoutedEventArgs e)
        {
            this.LogStackPanel.Children.Clear();
        }
    }
}