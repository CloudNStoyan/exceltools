using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Controls;
using ExcelTools.Attributes;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelTools.Pages
{
    [PageInfo(Header = "Convert to JSON")]
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

            string typeOfJsonConvert = this.TypeOfJsonConversionContainer.Children.OfType<RadioButton>()
                .First(r => r.IsChecked == true).DataContext.ToString();

            if (typeOfJsonConvert == "Rows")
            {
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
                        jsonObject.Add(propertyNames[j], propertyValues[j][i]);
                    }

                    jsonObjects.Add(jsonObject);
                }

                this.Output.OutputTextBox.Text = JsonConvert.SerializeObject(jsonObjects, Formatting.Indented);
            }
            else if (typeOfJsonConvert == "Columns")
            {
                var jsonProperties = (
                    from column in columns
                    let dataRows = excelWrapper.GetValueRows(ExcelWrapper.ConvertStringColumnToNumber(column))
                    select new JProperty(dataRows[0], dataRows.Skip(1))
                ).ToList();

                var jObject = new JObject(jsonProperties);

                this.Output.OutputTextBox.Text = JsonConvert.SerializeObject(jObject, Formatting.Indented);
            }
            else if (typeOfJsonConvert == "Excel")
            {
                var jsonData = (from column in columns
                        let columnNumber = ExcelWrapper.ConvertStringColumnToNumber(column)
                        select new JObject(new JProperty("Column", column), new JProperty("ColumnNumber", columnNumber),
                            new JProperty("Data", excelWrapper.GetValueRows(columnNumber)))
                    ).ToList();

                this.Output.OutputTextBox.Text = JsonConvert.SerializeObject(jsonData, Formatting.Indented);
            }
            else if (typeOfJsonConvert == "Array")
            {
                var jArray = new JArray();

                foreach (string column in columns)
                {
                    jArray.Add(excelWrapper.GetValueRows(ExcelWrapper.ConvertStringColumnToNumber(column)));
                }

                this.Output.OutputTextBox.Text = JsonConvert.SerializeObject(jArray, Formatting.Indented);
            }

            this.Output.FileName = excelWrapper.FileName.Split('.')[0] + ".json";
        }
    }
}