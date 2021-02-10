using System.Windows;
using System.Windows.Controls;
using ExcelTools.Attributes;

namespace ExcelTools.Pages
{
    [PageInfo(Header = "Excel to DataGrid")]
    public partial class ExcelToDataGrid : Page
    {
        public ExcelToDataGrid()
        {
            this.InitializeComponent();

            this.FileSelection.FileSelected += this.FileSelection_FileSelected;
        }

        private void FileSelection_FileSelected()
        {
            var excelWrapper = new ExcelWrapper(this.FileSelection.SelectedFile);

            string[] columns = excelWrapper.GetColumns();

            bool isFirst = true;

            foreach (string column in columns)
            {
                string[] values = excelWrapper.GetStringRows(ExcelWrapper.ConvertStringColumnToNumber(column));

                if (!isFirst)
                {
                    foreach (string value in values)
                    {
                        this.MainDataGrid.Items.Add(value);
                    }
                }
                else
                {
                    foreach (string value in values)
                    {
                        this.MainDataGrid.Columns.Add(new DataGridTextColumn { Header = value });
                    }
                    isFirst = false;
                }
            }

            this.MainDataGrid.Visibility = Visibility.Visible;
        }
    }
}