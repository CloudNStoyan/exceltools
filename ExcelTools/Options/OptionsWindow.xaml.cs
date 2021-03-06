﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
                this.RecentFilesListView.Items.Add(new TextBlock {Text = Path.GetFileName(filePath), DataContext = filePath, ToolTip = filePath});
            }

            string[][] multiFilePaths = SavedData.Config.RecentMultipleFiles;

            foreach (string[] multiFilePath in multiFilePaths)
            {
                this.RecentMultipleFilesListView.Items.Add(new TextBlock
                {
                    Text = string.Join(", ", multiFilePath.Select(Path.GetFileName)), DataContext = multiFilePath,
                    ToolTip = string.Join(", ", multiFilePath)
                });
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

        private void DeleteSelectedMultiplePaths(object sender, RoutedEventArgs e)
        {
            var items = this.RecentMultipleFilesListView.SelectedItems;

            var itemIds = new List<int>();

            for (int i = 0; i < items.Count; i++)
            {
                itemIds.Add(i);

                var textBlock = (TextBlock)items[i];

                SavedData.Config.RemoveFromRecentFiles((string[])textBlock.DataContext);
            }

            for (int i = 0; i < itemIds.Count; i++)
            {
                var id = itemIds[i];
                
                this.RecentMultipleFilesListView.Items.RemoveAt(id - i);
            }

            SavedData.Save();
        }
    }
}
