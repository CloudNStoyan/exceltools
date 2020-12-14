using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
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
                MessageBox.Show("The excel files must be different!");
                return;
            }

            string columnText = this.ColumnTextBox.Text;
            int column = ExcelWrapper.ConvertStringColumnToNumber(columnText);

            if (column == -1)
            {
                MessageBox.Show($"Column '{column}' is not a valid column!");
                return;
            }

            var selectedComparisonType = (ComparisonType)Enum.Parse(typeof(ComparisonType),this.ComparsionTypePanel.Children.OfType<RadioButton>()
                .First(r => r.IsChecked == true).DataContext.ToString());

            string[] firstExcelRows = firstExcelWrapper.GetStringRows(column).Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
            string[] secondExcelRows = secondExcelWrapper.GetStringRows(column).Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();

            if (selectedComparisonType == ComparisonType.FindSimilarities)
            {
                string[] logs = firstExcelRows.Where(x => secondExcelRows.Contains(x))
                    .Select(x => $"'{x}' was found in the both excels").ToArray();

                this.Logger.Log(logs.Length > 0 ? string.Join(Environment.NewLine, logs) : "No similarities were found between both excels!");
            }
            else if (selectedComparisonType == ComparisonType.FindDifferences)
            {
                string[] logs = firstExcelRows.Where(x => !secondExcelRows.Contains(x))
                    .Select(x => $"'{x}' was found in {firstExcelWrapper.FileName} but not in {secondExcelWrapper.FileName}").ToArray();

                this.Logger.Log(logs.Length > 0 ? string.Join(Environment.NewLine, logs) : "No differences were found between both excels!");
            }
        }

        private void CheckIfReady() => this.CompareButton.IsEnabled = this.FirstExcelFileSelected && this.SecondExcelFileSelected;

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            this.DataContext = this;
        }

        private void FirstFileSelection_OnFileChanged()
        {
            this.FirstExcelFileSelected = false;
            this.CheckIfReady();
        }

        private void FirstFileSelection_OnFileSelected()
        {
            this.FirstExcelFileSelected = true;
            this.CheckIfReady();
        }

        private void SecondFileSelection_OnFileChanged()
        {
            this.SecondExcelFileSelected = false;
            this.CheckIfReady();
        }

        private void SecondFileSelection_OnFileSelected()
        {
            this.SecondExcelFileSelected = true;
            this.CheckIfReady();
        }
    }
}