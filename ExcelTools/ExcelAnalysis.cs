﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;

namespace ExcelTools
{
    public class ExcelAnalysis
    {
        private Logger Logger { get; }
        public ExcelAnalysis(Logger logger)
        {
            this.Logger = logger;
        }

        public void MultipleFilesCountCells(string[] excelFiles)
        {
            int count = 0;

            foreach (string excelFile in excelFiles)
            {
                var excelWrapper = new ExcelWrapper(excelFile);
                count += excelWrapper.GetCount();
            }

            this.Logger.Log(count + " entries!");
        }

        public static int ConvertStringColumnToNumber(string column)
        {
            const string alphabet = "ABCDEFGHIJKLMNPPQRSTUVWX";

            return alphabet.IndexOf(column, StringComparison.Ordinal);
        }

        public void FindTool(ExcelWrapper excelWrapper, string column, string value, bool caseSensitive)
        {
            int columnNumber = ConvertStringColumnToNumber(column);

            if (columnNumber == -1)
            {
                MessageBox.Show($"Column '{column}' is not a valid column!");
                return;
            }

            string[] rows = excelWrapper.GetStringRows(columnNumber)?.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();

            if (rows == null)
            {
                MessageBox.Show($"There isn't {column} column in {excelWrapper.FileName}");
                return;
            }

            int index = 0;

            bool cellFound = false;

            for (int i = 0; i < rows.Length; i++)
            {
                string rowValue = rows[i];

                if (rowValue.Equals(value))
                {
                    index = i;
                    cellFound = true;
                } else if (!caseSensitive && rowValue.ToLower().Equals(value))
                {
                    index = i;
                    cellFound = true;
                }
            }

            this.Logger.Log(cellFound
                ? $"\"{value}\" was found at {column}{index + 1}"
                : $"\"{value}\" was not found!");
        }

        public void FindTool(ExcelWrapper[] excelWrappers, string column, string value, bool caseSensitive)
        {
            int columnNumber = ConvertStringColumnToNumber(column);

            if (columnNumber == -1)
            {
                MessageBox.Show($"Column '{column}' is not a valid column!");
                return;
            }

            var logs = new List<string>();
            foreach (var wrapper in excelWrappers)
            {
                string[] rows = wrapper.GetStringRows(columnNumber)?.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();

                if (rows == null)
                {
                    MessageBox.Show($"There isn't {column} column in {wrapper.FileName}");
                    return;
                }

                int index = 0;

                bool cellFound = false;

                for (int i = 0; i < rows.Length; i++)
                {
                    string rowValue = rows[i];

                    if (rowValue.Equals(value))
                    {
                        index = i;
                        cellFound = true;
                    }
                    else if (!caseSensitive && rowValue.ToLower().Equals(value))
                    {
                        index = i;
                        cellFound = true;
                    }
                }

                logs.Add(cellFound
                    ? $"\"{value}\" was found at {column}{index + 1} in \"{wrapper.FileName}\""
                    : $"\"{value}\" was not found!");
            }

            this.Logger.Log(string.Join("\r\n", logs));
        }

        public string[] ExportTool(ExcelWrapper excelWrapper, string column, bool skipEmpty = false)
        {
            int columnNumber = ConvertStringColumnToNumber(column);

            if (columnNumber == -1)
            {
                MessageBox.Show($"Column '{column}' is not a valid column!");
                return null;
            }

            string[] columnData = !skipEmpty ? excelWrapper.GetStringRows(columnNumber) : excelWrapper.GetStringRows(columnNumber)?.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();

            return columnData;
        }

        public string[] ExportTool(ExcelWrapper[] excelWrappers, string column, bool skipEmpty = false)
        {
            int columnNumber = ConvertStringColumnToNumber(column);

            if (columnNumber == -1)
            {
                MessageBox.Show($"Column '{column}' is not a valid column!");
                return null;
            }

            var list = new List<string>();

            foreach (var excelWrapper in excelWrappers)
            {
                string[] columnData = !skipEmpty ? excelWrapper.GetStringRows(columnNumber) : excelWrapper.GetStringRows(columnNumber)?.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();

                if (columnData == null)
                {
                    MessageBox.Show($"{excelWrapper.FileName} doesn't have '{column}' column!");
                    return null;
                }

                list.AddRange(columnData);
            }

            return list.ToArray();
        }

        public void FindDuplicates(ExcelWrapper excelWrapper, string column)
        {
            int columnNumber = ConvertStringColumnToNumber(column);

            if (columnNumber == -1)
            {
                MessageBox.Show($"Column '{column}' is not a valid column!");
                return;
            }

            string[] firstColumn = excelWrapper.GetStringRows(columnNumber).Where(x => x != null).ToArray();

            var dictionary = new Dictionary<string, int>();

            foreach (string cell in firstColumn)
            {
                if (!dictionary.ContainsKey(cell))
                {
                    dictionary.Add(cell, 1);
                }
                else
                {
                    dictionary[cell]++;
                }
            }

            var duplicates = new StringBuilder();

            bool duplicatesAreFound = false;

            foreach (var pair in dictionary)
            {
                if (pair.Value > 1)
                {
                    duplicatesAreFound = true;
                    duplicates.AppendLine($"'{pair.Key}' is entered {pair.Value} times");
                }
            }

            string output;

            if (duplicatesAreFound)
            {
                output = "Duplicates were found: \r\n" + duplicates;
            }
            else
            {
                output = "No duplicates found!";
            }

            this.Logger.Log(output.Trim());
        }
    }
}