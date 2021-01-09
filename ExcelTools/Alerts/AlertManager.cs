using System.Windows;

namespace ExcelTools.Alerts
{
    public static class AlertManager
    {
        public static void NoFileSelected()
        {
            var alert = new AlertWindow
            {
                Owner = Application.Current.MainWindow,
                AlertTitle = "No file selected",
                AlertBody = "You need to select file first!"
            };

            alert.ShowDialog();
        }
    }
}
