using System.Linq;
using System.Windows.Controls;
using ExcelTools.Attributes;

namespace ExcelTools.Pages
{
    [PageInfo(Header = "Convert to HTML")]
    public partial class ConvertToHtml : Page
    {
        public ConvertToHtml()
        {
            this.InitializeComponent();

            this.FileSelection.FileSelected += this.FileSelected;
        }

        private void FileSelected() => this.Convert(this.FileSelection.SelectedFile);

        private void Convert(string filePath)
        {
            var excelWrapper = new ExcelWrapper(filePath);

            string[] columns = excelWrapper.GetColumns();

            string typeOfJsonConvert = this.TypeOfHtmlConversionContainer.Children.OfType<RadioButton>()
                .First(r => r.IsChecked == true).DataContext.ToString();

            string[] propertyNames = columns.Select(column =>
                excelWrapper.GetValueRows(ExcelWrapper.ConvertStringColumnToNumber(column))[0]).ToArray();

            string[][] propertyValues = columns.Select(column =>
                    excelWrapper.GetValueRows(ExcelWrapper.ConvertStringColumnToNumber(column)).Skip(1).ToArray())
                .ToArray();

            string html = propertyNames.Aggregate("<table><tr>", (current, propertyName) => current + ("\r\n<th>" + propertyName + "</th>"));

            html += "</tr>";

            html = propertyValues.Aggregate(html, (current, values) => current + (string.Join("\r\n<tr>", values.Select(x => $"<td>{x}</td>")) + "</tr>"));

            html += "</table>";

            this.Output.OutputTextBox.Text = html;
            this.Output.FileName = excelWrapper.FileName.Split('.')[0] + ".html";
        }
    }
}
