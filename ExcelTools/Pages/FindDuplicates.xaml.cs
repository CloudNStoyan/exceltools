using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using ExcelTools.Attributes;

namespace ExcelTools.Pages
{
    [PageInfo(Header = "Duplicate Finder")]
    public partial class FindDuplicates : Page
    {
        private Logger Logger { get; }
        public FindDuplicates(Logger logger)
        {
            this.InitializeComponent();

            this.Logger = logger;
        }

        private void RunAnalysis(object sender, RoutedEventArgs e)
        {
            if (!File.Exists(this.FileSelection.SelectedFile))
            {
                MessageBox.Show("No file selected!");
                return;
            }

            var excelWrapper = new ExcelWrapper(this.FileSelection.SelectedFile);
            this.FindDuplicatesAnalysis(excelWrapper, this.ColumnTextBox.Text);
        }

        private void FindDuplicatesAnalysis(ExcelWrapper excelWrapper, string column)
        {
            int columnNumber = ExcelWrapper.ConvertStringColumnToNumber(column);

            if (columnNumber == -1)
            {
                MessageBox.Show($"Column '{column}' is not a valid column!");
                return;
            }

            string[] firstColumn = excelWrapper.GetStringRows(columnNumber).Where(x => x != null).ToArray();

            var dictionary = new Dictionary<string, int>();

            foreach (string cell in firstColumn)
            {
                if (!dictionary.ContainsKey(cell))
                {
                    dictionary.Add(cell, 1);
                }
                else
                {
                    dictionary[cell]++;
                }
            }

            var duplicates = new StringBuilder();

            bool duplicatesAreFound = false;

            foreach (var pair in dictionary)
            {
                if (pair.Value > 1)
                {
                    duplicatesAreFound = true;
                    duplicates.AppendLine($"'{pair.Key}' is entered {pair.Value} times");
                }
            }

            string output;

            if (duplicatesAreFound)
            {
                output = "Duplicates were found: \r\n" + duplicates;
            }
            else
            {
                output = "No duplicates found!";
            }

            this.Logger.Log(output.Trim());
        }
    }
}