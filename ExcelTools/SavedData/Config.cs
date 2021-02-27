using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;

namespace ExcelTools.SavedData
{
    public class Config
    {
        [JsonProperty]
        private List<string> recentFiles;
        [JsonIgnore]
        public string[] RecentFiles => this.recentFiles.ToArray();

        public Config()
        {
            this.recentFiles = new List<string>();
        }

        public void FillRecentFiles(string[] filePaths)
        {
            this.recentFiles.Clear();

            this.recentFiles.AddRange(filePaths);
        }

        public void AddToRecentFiles(string filePath)
        {
            var oldPaths = this.recentFiles.Take(4);

            this.recentFiles = new List<string> {filePath};

            this.recentFiles.AddRange(oldPaths);
        }
    }
}
