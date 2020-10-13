using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelTools
{
    public class ExcelAnalysis
    {
        private Logger Logger { get; }
        public ExcelAnalysis(Logger logger)
        {
            this.Logger = logger;
        }

        public void MultipleFilesCountCells(string folderPath)
        {
            string[] excelFiles = Directory.GetFiles(folderPath);

            int count = 0;

            foreach (string excelFile in excelFiles)
            {
                var excelWrapper = new ExcelWrapper(excelFile);
                count += excelWrapper.GetStringRows(0).Where(x => x != null).ToArray().Length;
            }

            this.Logger.Log("Count: " + count);
        }

        private int ConvertStringColumnToNumber(string column)
        {
            const string alphabet = "ABCDEFGHIJKLMNPPQRSTUVWX";

            return alphabet.IndexOf(column, StringComparison.Ordinal);
        }

        public void FindDuplicates(ExcelWrapper excelWrapper, string column)
        {
            int columnNumber = this.ConvertStringColumnToNumber(column);

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