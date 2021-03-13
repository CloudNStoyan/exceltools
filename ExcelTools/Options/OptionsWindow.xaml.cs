using System.Windows;
using System.Windows.Controls;
using ExcelTools.DataSaving;

namespace ExcelTools.Options
{
    public partial class OptionsWindow : Window
    {
        public OptionsWindow()
        {
            this.InitializeComponent();

            string[] filePaths = SavedData.Config.RecentFiles;

            foreach (string filePath in filePaths)
            {
                this.RecentFilesListView.Items.Add(new TextBlock {Text = filePath});
            }
        }

        private void DeleteSelectedPaths(object sender, RoutedEventArgs e)
        {
            var items = this.RecentFilesListView.SelectedItems;

            foreach (var item in items)
            {
                var textBlock = (TextBlock) item;

                SavedData.Config.RemoveFromRecentFiles(textBlock.Text);
            }

            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];

                this.RecentFilesListView.Items.Remove(item);
            }

            SavedData.Save();
        }
    }
}
