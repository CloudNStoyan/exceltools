using System.Collections;
using System.Collections.Generic;
using System.IO;
using ExcelDataReader;

namespace ExcelTools
{
    public class ExcelWrapper
    {
        private string FilePath { get; }
        public ExcelWrapper(string filePath)
        {
            this.FilePath = filePath;
        }

        public string[] GetCountByColor(string color)
        {
            using (var stream = File.Open(this.FilePath, FileMode.Open, FileAccess.Read))
            {
                var a = new List<string>();

                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var data = reader.GetData(0);
                    foreach (var extendedProperty in data.GetSchemaTable().Columns[0].ExtendedProperties)
                    {
                        a.Add(extendedProperty.ToString());
                    }
                }

                return a.ToArray();
            }
        }

        public int GetCount()
        {
            using (var stream = File.Open(this.FilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    return reader.RowCount;
                }
            }
        }

        public string[] GetStringRows(int col)
        {
            using (var stream = File.Open(this.FilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var rowsData = new List<string>();

                    do
                    {
                        while (reader.Read())
                        {
                            var fieldType = reader.GetFieldType(col);

                            rowsData.Add(fieldType == typeof(string) ? reader.GetString(col) : null);
                        }
                    } while (reader.NextResult());

                    return rowsData.ToArray();
                }
            }
        }

        public double[] GetDoubleRows(int col)
        {
            using (var stream = File.Open(this.FilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var rowsData = new List<double>();

                    do
                    {
                        while (reader.Read())
                        {
                            var fieldType = reader.GetFieldType(col);
                            
                            rowsData.Add(fieldType == typeof(double) ? reader.GetDouble(col) : -1);
                        }
                    } while (reader.NextResult());

                    return rowsData.ToArray();
                }
            }
        }
    }
}
