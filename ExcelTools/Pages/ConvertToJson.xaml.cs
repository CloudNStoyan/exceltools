using System.IO;
using System.Linq;
using System.Windows.Controls;
using ExcelTools.Attributes;
using Newtonsoft.Json;

namespace ExcelTools.Pages
{
    [PageInfo(Header = "Convert To Json")]
    public partial class ConvertToJson : Page
    {
        public ConvertToJson()
        {
            this.InitializeComponent();

            this.FileSelection.FileSelected += this.Convert;
        }

        private void Convert()
        {
            if (!this.FileSelection.FileIsSelected)
            {
                AlertManager.NoFileSelected();
                return;
            }

            if (!File.Exists(this.FileSelection.SelectedFile))
            {
                AlertManager.Custom("The file doesn't exist!");
                return;
            }

            var excelWrapper = new ExcelWrapper(this.FileSelection.SelectedFile);

            string[] columns = excelWrapper.GetColumns();

            var jsonData = (from column in columns
                    let columnNumber = ExcelWrapper.ConvertStringColumnToNumber(column)
                    select new ExcelJsonData
                        {Column = column, ColumnNumber = columnNumber, Data = excelWrapper.GetValueRows(columnNumber)})
                .ToList();

            this.Output.FileName = excelWrapper.FileName.Split('.')[0] + ".json";

            this.Output.OutputTextBox.Text = JsonConvert.SerializeObject(jsonData, Formatting.Indented);

            this.FileSelection.ChangeFileHandler(null, null);
        }

        public class ExcelJsonData
        {
            public string Column { get; set; }
            public int ColumnNumber { get; set; }
            public string[] Data { get; set; }
        }
    }
}
