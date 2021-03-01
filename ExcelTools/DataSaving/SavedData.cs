using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelTools.DataSaving
{
    public static class SavedData
    {
        private const string ConfigPath = "./config.json";
        public static Config Config { get; private set; }

        public static void LoadData()
        {
            if (File.Exists(ConfigPath))
            {
                string json = File.ReadAllText(ConfigPath);

                if (ValidateJSON(json))
                {
                    Config = JsonConvert.DeserializeObject<Config>(json);
                }
            }

            if (Config == null)
            {
                Config = new Config();
            }
        }

        public static void Save() => File.WriteAllText(ConfigPath, JsonConvert.SerializeObject(Config));

        private static bool ValidateJSON(string json)
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
