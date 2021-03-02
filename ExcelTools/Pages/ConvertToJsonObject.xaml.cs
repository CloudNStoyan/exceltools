using System.IO;
using System.Linq;
using System.Windows.Controls;
using ExcelTools.Attributes;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelTools.Pages
{
    [PageInfo(Header = "Convert to JSON Object")]
    public partial class ConvertToJsonObject : Page
    {
        public ConvertToJsonObject()
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

            var jsonProperties = (
                from column in columns
                let dataRows = excelWrapper.GetValueRows(ExcelWrapper.ConvertStringColumnToNumber(column))
                let obj = new JObject()
                select new JProperty(dataRows[0], dataRows.Skip(1))
                ).ToList();

            var jObject = new JObject(jsonProperties);

            this.Output.FileName = excelWrapper.FileName.Split('.')[0] + ".json";
            this.Output.OutputTextBox.Text = JsonConvert.SerializeObject(jObject, Formatting.Indented);
        }
    }
}
