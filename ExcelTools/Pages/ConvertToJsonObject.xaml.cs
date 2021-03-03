using System.Collections.Generic;
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

            string[] propertyNames = columns.Select(column =>
                excelWrapper.GetValueRows(ExcelWrapper.ConvertStringColumnToNumber(column))[0]).ToArray();

            string[][] propertyValues = columns.Select(column =>
                    excelWrapper.GetValueRows(ExcelWrapper.ConvertStringColumnToNumber(column)).Skip(1).ToArray())
                .ToArray();

            int jObjectCount = propertyValues.Select(propertyValue => propertyValue.Length).Max();

            var jsonObjects = new List<JObject>();

            for (int i = 0; i < jObjectCount; i++)
            {
                var jsonObject = new JObject();

                for (int j = 0; j < propertyNames.Length; j++)
                {
                    jsonObject.Add(propertyNames[j],propertyValues[j][i]);
                }

                jsonObjects.Add(jsonObject);
            }


            //var jsonProperties = (
            //    from column in columns
            //    let dataRows = excelWrapper.GetValueRows(ExcelWrapper.ConvertStringColumnToNumber(column))
            //    let obj = new JObject()
            //    select new JProperty(dataRows[0], dataRows.Skip(1))
            //    ).ToList();
            //
            //var jObject = new JObject(jsonProperties);
            //
            //this.Output.FileName = excelWrapper.FileName.Split('.')[0] + ".json";
            //this.Output.OutputTextBox.Text = JsonConvert.SerializeObject(jObject, Formatting.Indented);
            this.Output.FileName = excelWrapper.FileName.Split('.')[0] + ".json";
            this.Output.OutputTextBox.Text = JsonConvert.SerializeObject(jsonObjects, Formatting.Indented);

        }
    }
}
