using System.Collections.Generic;
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

            var itemIds = new List<int>();

            for (int i = 0; i < items.Count; i++)
            {
                itemIds.Add(i);

                var textBlock = (TextBlock)items[i];

                SavedData.Config.RemoveFromRecentFiles(textBlock.Text);
            }

            for (int i = 0; i < itemIds.Count; i++)
            {
                var id = itemIds[i];

                this.RecentFilesListView.Items.RemoveAt(id - i);
            }

            SavedData.Save();
        }
    }
}
