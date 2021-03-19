using System.Collections.Generic;
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

            string html = "<table>\r\n\t<tbody>\r\n\t\t<tr>";

            html = columns
                .Select(column =>
                excelWrapper.GetValueRows(ExcelWrapper.ConvertStringColumnToNumber(column))[0])
                .Aggregate(html, (current, propertyName) => current + $"\r\n\t\t\t<th>{propertyName}</th>");

            html += "\r\n\t\t</tr>";
            
            string[][] propertyValues = columns.Select(column =>
                    excelWrapper.GetValueRows(ExcelWrapper.ConvertStringColumnToNumber(column)).Skip(1).ToArray())
                .ToArray();

            var list = new List<List<string>>();

            foreach (string[] values in propertyValues)
            {
                for (int j = 0; j < values.Length; j++)
                {
                    if (list.Count >= j + 1)
                    {
                        list[j].Add(values[j]);
                    }
                    else
                    {
                        list.Add(new List<string>{values[j]});
                    }
                }
            }

            html = list.Aggregate(html,
                       (current, row) =>
                           current +
                           $"\r\n\t\t<tr>\r\n{string.Join("\r\n", row.Select(x => $"\t\t\t<td>{x}</td>"))}\r\n\t\t</tr>") +
                   "\r\n\t</tbody>\r\n</table>";

            this.Output.OutputTextBox.Text = html;
            this.Output.FileName = excelWrapper.FileName.Split('.')[0] + ".html";
        }
    }
}
