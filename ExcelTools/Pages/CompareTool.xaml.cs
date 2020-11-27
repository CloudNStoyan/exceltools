using System.Collections.Generic;
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
        public CompareTool(Logger logger)
        {
            this.InitializeComponent();

            this.Logger = logger;
        }

        private void OnFileSelected()
        {
            this.CompareButton.IsEnabled = true;
        }

        private void OnFileChanged()
        {
            this.CompareButton.IsEnabled = false;
        }

        private void RunAnalysis(object sender, RoutedEventArgs e)
        {
            string[] filePaths = {""};

            if (filePaths.Length < 2)
            {
                return;
            }

            var excelWrappers = filePaths.Select(x => new ExcelWrapper(x)).ToArray();

            string columnText = this.ColumnTextBox.Text;
            int column = ExcelAnalysis.ConvertStringColumnToNumber(columnText);

            string[] firstExcelRows = excelWrappers[0].GetStringRows(column);
            string[] secondExcelRows = excelWrappers[1].GetStringRows(column);

            foreach (string rowValue in firstExcelRows)
            {
                if (secondExcelRows.Contains(rowValue))
                {
                    this.Logger.Log($"'{rowValue}' was found in the two excels");
                }
            }

        }
    }
}
