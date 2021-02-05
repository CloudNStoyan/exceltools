using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace ExcelTools.SavedData
{
    public class SavedData
    {
        private const string ConfigPath = "./config.json";
        public Config Config { get; private set; }
        public SavedData()
        {
            this.LoadData();
        }

        private void LoadData()
        {
            this.Config = File.Exists(ConfigPath) ? JsonConvert.DeserializeObject<Config>(File.ReadAllText(ConfigPath)) : new Config();
        }

        public void Save() => File.WriteAllText(ConfigPath, JsonConvert.SerializeObject(this.Config));
    }
}
