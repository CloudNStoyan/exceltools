using System.Collections.Generic;
using System.Linq;
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
            if (!this.FileSelection.FileIsSelected)
            {
                AlertManager.NoFileSelected();
                return;
            }

            var excelWrappers = this.FileSelection.MultipleFilesChecked
                ? this.FileSelection.SelectedFiles.Select(x => new ExcelWrapper(x)).ToArray()
                : new[] {new ExcelWrapper(this.FileSelection.SelectedFile)};

            this.FindDuplicatesAnalysis(excelWrappers, this.SpecificColumnCheckBox.IsChecked == true);
        }

        private void FindDuplicatesAnalysis(ExcelWrapper[] excelWrappers, bool specificColumns)
        {
            foreach (var excelWrapper in excelWrappers)
            {
                var duplicates = new List<string>();

                string[] columns = specificColumns ? this.ColumnTextBox.GetColumns() : excelWrapper.GetColumns();

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

                        string checkingColumn = $"Checking column '{column}'";

                        duplicates.Add(checkingColumn);
                        duplicates.AddRange(from pair in cellEntries where pair.Value > 1 select $"'{pair.Key}' is entered {pair.Value} times");

                        if (duplicates.Last() == checkingColumn)
                        {
                            duplicates.Add("Nothing found!");
                        }

                    } else
                    {
                        duplicates.Add($"There is no '{column}' column in {excelWrapper.FileName}");
                    }
                }

                string log = string.Join("\n", duplicates);

                this.Logger.Log($"{excelWrapper.FileName}\r\n" + log);
            }
        }
    }
}