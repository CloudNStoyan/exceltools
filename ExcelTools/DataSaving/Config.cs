using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;

namespace ExcelTools.DataSaving
{
    public class Config
    {
        [JsonProperty] private List<string> recentFiles = new List<string>();

        [JsonIgnore] public string[] RecentFiles => this.recentFiles.ToArray();

        public void AddToRecentFiles(string filePath)
        {
            this.recentFiles.RemoveAll(x => x.Equals(filePath));

            var oldPaths = this.recentFiles.Take(4);

            this.recentFiles = new List<string> {filePath};

            this.recentFiles.AddRange(oldPaths);
        }

        public void RemoveFromRecentFiles(string filePath) => this.recentFiles.Remove(filePath);

        [JsonProperty] private List<string[]> recentMultipleFiles = new List<string[]>();

        [JsonIgnore] public string[][] RecentMultipleFiles => this.recentMultipleFiles.ToArray();

        public void AddToRecentFiles(string[] filePaths)
        {
            this.recentMultipleFiles.RemoveAll(x => AreArraysEqual(x, filePaths));

            var oldPaths = this.recentMultipleFiles.Take(4);

            this.recentMultipleFiles = new List<string[]> {filePaths};

            this.recentMultipleFiles.AddRange(oldPaths);
        }

        public void RemoveFromRecentFiles(string[] filePaths) => this.recentMultipleFiles.Remove(filePaths);

        private static bool AreArraysEqual(string[] firstArray, string[] secondArray)
        {
            if (firstArray.Length != secondArray.Length)
            {
                return false;
            }

            return !firstArray.Where((t, i) => t != secondArray[i]).Any();
        }
    }
}