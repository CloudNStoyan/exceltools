using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
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
        }

        private void Convert(object sender, RoutedEventArgs e)
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

            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                excelWrapper.FileName.Split('.')[0] + ".json");

            File.WriteAllText(path, JsonConvert.SerializeObject(jsonData));

            AlertManager.Custom("Success!");
        }

        public class ExcelJsonData
        {
            public string Column { get; set; }
            public int ColumnNumber { get; set; }
            public string[] Data { get; set; }
        }
    }
}
