using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

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
            if (File.Exists(ConfigPath))
            {
                string json = File.ReadAllText(ConfigPath);

                if (this.ValidateJSON(json))
                {
                    this.Config = JsonConvert.DeserializeObject<Config>(json);
                }
            }

            if (this.Config == null)
            {
                this.Config = new Config();
            }
        }

        public void Save() => File.WriteAllText(ConfigPath, JsonConvert.SerializeObject(this.Config));

        private bool ValidateJSON(string json)
        {
            try
            {
                JToken.Parse(json);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
