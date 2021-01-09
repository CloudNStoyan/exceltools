using System.Windows;

namespace ExcelTools.Alerts
{
    public partial class AlertWindow : Window
    {
        public string AlertTitle { get; set; }
        public string AlertBody { get; set; }
        public AlertWindow()
        {
            this.InitializeComponent();
        }

        private void AlertWindow_OnLoaded(object sender, RoutedEventArgs e)
        {
            this.DataContext = this;
        }

        private void OkHandler(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
