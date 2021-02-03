using System.IO;

namespace ExcelTools
{
    public static class SavedData
    {
        private const string SavedDataPath = "./savedData.txt";
        public static string[] RecentFiles
        {
            get => File.Exists(SavedDataPath) ? File.ReadAllLines(SavedDataPath) : null;
            set => File.WriteAllLines(SavedDataPath, value);
        }
    }
}
