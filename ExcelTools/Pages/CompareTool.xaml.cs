﻿using System.Linq;
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

        private void RunAnalysis(object sender, RoutedEventArgs e)
        {
            var firstExcelWrapper = new ExcelWrapper(this.FirstFileSelection.SelectedFile);
            var secondExcelWrapper = new ExcelWrapper(this.SecondFileSelection.SelectedFile);

            string columnText = this.ColumnTextBox.Text;
            int column = ExcelAnalysis.ConvertStringColumnToNumber(columnText);

            if (column == -1)
            {
                MessageBox.Show($"Column '{column}' is not a valid column!");
                return;
            }

            string[] firstExcelRows = firstExcelWrapper.GetStringRows(column);
            string[] secondExcelRows = secondExcelWrapper.GetStringRows(column);

            foreach (string rowValue in firstExcelRows)
            {
                if (secondExcelRows.Contains(rowValue))
                {
                    this.Logger.Log($"'{rowValue}' was found in the two excels");
                }
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