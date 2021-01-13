using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ExcelTools.Alerts;
using ExcelTools.Attributes;

namespace ExcelTools.Pages
{
    [PageInfo(Header = "Compare Excels")]
    public partial class CompareTool : Page
    {
        private Logger Logger { get; }
        private bool FirstExcelFileSelected { get; set; }
        private bool SecondExcelFileSelected { get; set; }
        public CompareTool(Logger logger)
        {
            this.InitializeComponent();

            this.Logger = logger;
        }

        private enum ComparisonType
        {
            FindSimilarities,
            FindDifferences
        }

        private void RunAnalysis(object sender, RoutedEventArgs e)
        {
            var firstExcelWrapper = new ExcelWrapper(this.FirstFileSelection.SelectedFile);
            var secondExcelWrapper = new ExcelWrapper(this.SecondFileSelection.SelectedFile);

            if (firstExcelWrapper.FileName == secondExcelWrapper.FileName)
            {
                AlertManager.Custom("The excel files must be different!");
                return;
            }

            string[] columns = this.ColumnTextBox.GetColumns();

            var selectedComparisonType = (ComparisonType)Enum.Parse(typeof(ComparisonType), this.ComparisonTypePanel.Children.OfType<RadioButton>()
                .First(r => r.IsChecked == true).DataContext.ToString());

            switch (selectedComparisonType)
            {
                case ComparisonType.FindSimilarities:
                {
                    var logs = new List<string>();

                    foreach (string column in columns)
                    {
                        int columnNumber = ExcelWrapper.ConvertStringColumnToNumber(column);

                        string[] firstExcelRows = firstExcelWrapper.GetStringRows(columnNumber)
                            .Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
                        string[] secondExcelRows = secondExcelWrapper.GetStringRows(columnNumber)
                            .Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
                        logs.AddRange(firstExcelRows.Where(x => secondExcelRows.Contains(x))
                            .Select(x => $"'{x}' was found in the both excels"));
                    }

                    this.Logger.Log(logs.Count > 0
                        ? string.Join(Environment.NewLine, logs)
                        : "No similarities were found between both excels!");

                    break;
                }
                case ComparisonType.FindDifferences:
                {
                    var logs = new List<string>();

                    foreach (string column in columns)
                    {
                        int columnNumber = ExcelWrapper.ConvertStringColumnToNumber(column);

                        string[] firstExcelRows = firstExcelWrapper.GetStringRows(columnNumber)
                            .Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
                        string[] secondExcelRows = secondExcelWrapper.GetStringRows(columnNumber)
                            .Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
                        logs.AddRange(firstExcelRows.Where(x => !secondExcelRows.Contains(x))
                            .Select(x => $"'{x}' was found in {firstExcelWrapper.FileName} but not in {secondExcelWrapper.FileName}"));
                    }

                    this.Logger.Log(logs.Count > 0
                        ? string.Join(Environment.NewLine, logs)
                        : "No differences were found between both excels!");

                    break;
                }
            }
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            this.DataContext = this;
        }
    }
}