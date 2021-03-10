using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ExcelTools.Attributes;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelTools.Pages
{
    [PageInfo(Header = "Convert to JSON Array")]
    public partial class ConvertToJsonArray : Page
    {
        private string LastPath { get; set; }

        public ConvertToJsonArray()
        {
            this.InitializeComponent(); 

            this.FileSelection.FileSelected += this.ConvertHandler;

            this.InputColumn.InputChanged += this.ReselectFile;

            this.InputRow.NumberChanged += this.ReselectFile;
        }

        private void ConvertHandler()
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

            string selectedFile = this.FileSelection.SelectedFile;

            this.Convert(selectedFile);

            this.LastPath = selectedFile;
        }

        private void Convert(string filePath)
        {
            var excelWrapper = new ExcelWrapper(filePath);

            string typeOfJsonConvert = this.TypeOfJsonConversionContainer.Children.OfType<RadioButton>()
                .First(r => r.IsChecked == true).DataContext.ToString();

            object json = null;

            switch (typeOfJsonConvert)
            {
                case "Rows":
                {
                    json = RowConverter(excelWrapper, this.InputRow.Number - 1);
                    break;
                }
                case "Columns":
                {
                    json = ColumnConverted(excelWrapper, this.InputColumn.Text);
                    break;
                }
                case "Excel":
                {
                    json = ExcelConverter(excelWrapper);
                    break;
                }
            }


            this.Output.OutputTextBox.Text = JsonConvert.SerializeObject(json, Formatting.Indented);
            this.Output.FileName = excelWrapper.FileName.Split('.')[0] + ".json";
        }

        private static object ExcelConverter(ExcelWrapper excelWrapper)
        {
            string[] columns = excelWrapper.GetColumns();

            var jArray = new JArray();

            foreach (string column in columns)
            {
                // ReSharper disable once CoVariantArrayConversion
                jArray.Add(new JArray(excelWrapper.GetValueRows(ExcelWrapper.ConvertStringColumnToNumber(column))));
            }

            return jArray;
        }

        private static object ColumnConverted(ExcelWrapper excelWrapper, string column)
        {
            // ReSharper disable once CoVariantArrayConversion
            var jArray = new JArray(excelWrapper.GetValueRows(ExcelWrapper.ConvertStringColumnToNumber(column)));

            return jArray;
        }

        private static object RowConverter(ExcelWrapper excelWrapper, int row)
        {
            string[] columns = excelWrapper.GetColumns();

            var jArray = new JArray();

            foreach (string column in columns)
            {
                jArray.Add(excelWrapper.GetValueRows(ExcelWrapper.ConvertStringColumnToNumber(column))[row]);
            }

            return jArray;
        }

        private void RadioButtonChecked(object sender, RoutedEventArgs e) => this.ReselectFile();

        private void ReselectFile()
        {
            if (this.LastPath == null)
            {
                return;
            }

            this.Convert(this.LastPath);
        }
    }
}
