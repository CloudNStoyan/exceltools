using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using ExcelTools.Alerts;
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
            AlertManager.SetupAlert(this.AlertBox,this.AlertBody);

            this.SetupPages();
        }

        private Button ActiveButton { get; set; }

        private void SetupPages()
        {
            var pageClasses = Assembly.GetExecutingAssembly().DefinedTypes.Where(x => x.Namespace == "ExcelTools.Pages" && x.BaseType?.Name == "Page").ToArray();

            var buttons = new List<Button>();

            foreach (var pageClass in pageClasses)
            {
                bool hasLoggerParameter = false;

                var constructorParameters = pageClass.GetConstructors().Select(x => x.GetParameters());

                foreach (var parameters in constructorParameters)
                {
                    foreach (var parameterInfo in parameters)
                    {
                        if (parameterInfo.Name == "logger")
                        {
                            hasLoggerParameter = true;
                        }
                    }
                }

                var instance = hasLoggerParameter ? (Page) Activator.CreateInstance(pageClass, this.Logger) : (Page) Activator.CreateInstance(pageClass);

                var pageInfo = pageClass.GetCustomAttribute<PageInfo>();

                var button = new Button
                {
                    Content = pageInfo.Header,
                    Padding = new Thickness(5),
                    Margin = new Thickness(5),
                    ToolTip = $"Navigate to {pageInfo.Header}"
                };

                button.Click += (sender, args) => this.NavigateToPage(instance, button);

                button.DataContext = pageInfo.Order;
                
                buttons.Add(button);
                
                if (pageInfo.Order == 0 && pageInfo.ShowInNavigation)
                {
                    this.NavigateToPage(instance, button);
                }
            }

            buttons = buttons.OrderBy(x => int.Parse(x.DataContext.ToString())).ToList();

            foreach (var button in buttons)
            {
                this.NavigationContainer.Children.Add(button);
            }

        }

        private void NavigateToPage(Page page, Button button)
        {
            if (this.ActiveButton != null)
            {
                this.ActiveButton.Background = SystemColors.ControlLightBrush;
            }

            this.ActiveButton = button;
            this.ActiveButton.Background = SystemColors.ControlDarkBrush;

            this.Settings.Navigate(page);
        }

        private void ClearLogStackPanelHandler(object sender, RoutedEventArgs e)
        {
            this.LogStackPanel.Children.Clear();
        }

        private void CloseAlert(object sender, RoutedEventArgs e)
        {
            this.AlertBox.Visibility = Visibility.Hidden;
        }
    }
}