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

        private void RunAnalysis(object sender, RoutedEventArgs e)
        {
            var firstExcelWrapper = new ExcelWrapper(this.FirstFileSelection.SelectedFile);
            var secondExcelWrapper = new ExcelWrapper(this.SecondFileSelection.SelectedFile);

            string columnText = this.ColumnTextBox.Text;
            int column = ExcelAnalysis.ConvertStringColumnToNumber(columnText);

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
    }
}
