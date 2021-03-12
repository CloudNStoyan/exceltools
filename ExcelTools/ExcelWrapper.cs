using System;
using System.Collections.Generic;
using System.IO;
using ExcelDataReader;

namespace ExcelTools
{
    public class ExcelWrapper
    {
        private string FilePath { get; }
        public string FileName { get; }
        private string Extension { get; }
        private bool IsCsv { get; }
        public ExcelWrapper(string filePath)
        {
            this.FilePath = filePath;
            this.FileName = Path.GetFileName(filePath);
            this.Extension = Path.GetExtension(filePath);
            this.IsCsv = this.Extension == ".csv";
        }

        public static int ConvertStringColumnToNumber(string column)
        {
            const string alphabet = "ABCDEFGHIJKLMNPPQRSTUVWX";

            return alphabet.IndexOf(column, StringComparison.Ordinal);
        }

        private static string ConvertStringColumnToNumber(int column)
        {
            const string alphabet = "ABCDEFGHIJKLMNPPQRSTUVWX";

            return alphabet.Length > column ? alphabet[column].ToString() : null;
        }

        public int GetCount()
        {
            using (var stream = File.Open(this.FilePath, FileMode.Open, FileAccess.Read))
            {
                if (!this.IsCsv)
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        return reader.RowCount;
                    }
                }

                using (var reader = ExcelReaderFactory.CreateCsvReader(stream))
                {
                    return reader.RowCount;
                }
            }
        }

        public string[] GetColumns()
        {
            var columns = new List<string>();
            using (var stream = File.Open(this.FilePath, FileMode.Open, FileAccess.Read))
            {
                if (!this.IsCsv)
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            columns.Add(ConvertStringColumnToNumber(i));
                        }
                    }
                }
                else
                {
                    using (var reader = ExcelReaderFactory.CreateCsvReader(stream))
                    {
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            columns.Add(ConvertStringColumnToNumber(i));
                        }
                    }
                }
            }

            return columns.ToArray();
        }

        public string[] GetValueRows(int col)
        {
            using (var stream = File.Open(this.FilePath, FileMode.Open, FileAccess.Read))
            {
                var rowsData = new List<string>();

                if (!this.IsCsv)
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        do
                        {
                            while (reader.Read())
                            {
                                if (col >= reader.FieldCount)
                                {
                                    return null;
                                };

                                rowsData.Add(reader.GetValue(col)?.ToString());
                            }
                        } while (reader.NextResult());
                    }
                }
                else
                {
                    using (var reader = ExcelReaderFactory.CreateCsvReader(stream))
                    {
                        do
                        {
                            while (reader.Read())
                            {
                                if (col >= reader.FieldCount)
                                {
                                    return null;
                                };

                                rowsData.Add(reader.GetValue(col)?.ToString());
                            }
                        } while (reader.NextResult());
                    }
                }

                return rowsData.ToArray();
            }
        }

        public string[] GetStringRows(int col)
        {
            using (var stream = File.Open(this.FilePath, FileMode.Open, FileAccess.Read))
            {
                var rowsData = new List<string>();

                if (!this.IsCsv)
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        do
                        {
                            while (reader.Read())
                            {
                                if (col >= reader.FieldCount)
                                {
                                    return null;
                                };

                                var fieldType = reader.GetFieldType(col);

                                rowsData.Add(fieldType == typeof(string) ? reader.GetString(col) : null);
                            }
                        } while (reader.NextResult());
                    }
                }
                else
                {
                    using (var reader = ExcelReaderFactory.CreateCsvReader(stream))
                    {
                        do
                        {
                            while (reader.Read())
                            {
                                if (col >= reader.FieldCount)
                                {
                                    return null;
                                };

                                var fieldType = reader.GetFieldType(col);

                                rowsData.Add(fieldType == typeof(string) ? reader.GetString(col) : null);
                            }
                        } while (reader.NextResult());
                    }
                }

                return rowsData.ToArray();
            }
        }

        public double[] GetDoubleRows(int col)
        {
            using (var stream = File.Open(this.FilePath, FileMode.Open, FileAccess.Read))
            {
                var rowsData = new List<double>();

                if (!this.IsCsv)
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        do
                        {
                            while (reader.Read())
                            {
                                if (col >= reader.FieldCount)
                                {
                                    return null;
                                };

                                var fieldType = reader.GetFieldType(col);

                                rowsData.Add(fieldType == typeof(double) ? reader.GetDouble(col) : -1);
                            }
                        } while (reader.NextResult());
                    }
                }
                else
                {
                    using (var reader = ExcelReaderFactory.CreateCsvReader(stream))
                    {
                        do
                        {
                            while (reader.Read())
                            {
                                if (col >= reader.FieldCount)
                                {
                                    return null;
                                };

                                if (double.TryParse(reader.GetString(col), out double result))
                                {
                                    rowsData.Add(result);
                                }
                                else
                                {
                                    rowsData.Add(-1);
                                }
                            }
                        } while (reader.NextResult());
                    }
                }

                return rowsData.ToArray();
            }
        }
    }
}
