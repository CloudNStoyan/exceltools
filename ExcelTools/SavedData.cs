using System.Collections.Generic;
using System.IO;

namespace ExcelTools
{
    public class SavedData
    {
        private const string SavedDataPath = "./savedData.txt";
        public string[] RecentFiles { get; set; }

        public SavedData()
        {
            this.LoadSavedData();
        }

        public void AddFilePath(string filePath)
        {
            var recentFiles = new List<string>();

            if (this.RecentFiles != null)
            {
                recentFiles.AddRange(this.RecentFiles);
            }

            recentFiles.Add(filePath);

            this.RecentFiles = recentFiles.ToArray();

            File.WriteAllLines(SavedDataPath, this.RecentFiles);
        }

        private void LoadSavedData()
        {
            if (File.Exists(SavedDataPath))
            {
                this.RecentFiles = File.ReadAllLines(SavedDataPath);
            }
        }
    }
}
