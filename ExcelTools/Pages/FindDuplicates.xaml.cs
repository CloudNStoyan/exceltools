using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ExcelTools.Alerts;
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
                AlertManager.NoFileSelected();
                return;
            }

            var excelWrapper = new ExcelWrapper(this.FileSelection.SelectedFile);

            string[] columns = this.ColumnTextBox.GetColumns();

            this.FindDuplicatesAnalysis(excelWrapper,
                this.SpecificColumnCheckBox.IsChecked == true ? columns : excelWrapper.GetColumns());
        }

        private void FindDuplicatesAnalysis(ExcelWrapper excelWrapper, string[] columns)
        {
            var duplicates = new List<string>();

            foreach (string column in columns)
            {
                int columnNumber = ExcelWrapper.ConvertStringColumnToNumber(column);

                string[] cells = excelWrapper.GetStringRows(columnNumber)?.Where(cell => !string.IsNullOrEmpty(cell)).ToArray();

                if (cells != null)
                {
                    var cellEntries = new Dictionary<string, int>();

                    foreach (string cell in cells)
                    {
                        if (!cellEntries.ContainsKey(cell))
                        {
                            cellEntries.Add(cell, 1);
                        }
                        else
                        {
                            cellEntries[cell]++;
                        }
                    }

                    duplicates.Add($"Checking column '{column}'");
                    duplicates.AddRange(from pair in cellEntries where pair.Value > 1 select $"'{pair.Key}' is entered {pair.Value} times");
                } else
                {
                    duplicates.Add($"There is no '{column}' column in {excelWrapper.FileName}");
                }
            }

            this.Logger.Log(duplicates.Count > 0 ? "Duplicates were found: \n" + string.Join("\n", duplicates) : "No duplicates found!");
        }
    }
}