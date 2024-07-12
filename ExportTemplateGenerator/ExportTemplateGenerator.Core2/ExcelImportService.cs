
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExportTemplateGenerator.Core2
{
    public class ExcelImportService
    {
        private string _fileName = null;
        private string _sheetName = null;
        public ExcelImportService(string fileName)
        {
            _fileName = fileName;
        }

        public void GetData()
        {
            //var excelApp = new Application();
            //var excelWorkbook = excelApp.Workbooks.Open(_fileName, 0, false, 5, "", "", false, 
            //    Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //var excelSheets = excelWorkbook.Worksheets;
            //Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = 
            //    (Microsoft.Office.Interop.Excel.Worksheet)excelSheets.get_Item(_sheetName);
            //var column1 = excelWorksheet.Columns[0];
            //var column2 = excelWorksheet.Columns[1];
            var dictionary = new Dictionary<string, string>();
            const Int32 BufferSize = 128;
            using (var fileStream = File.OpenRead(_fileName))
            using (var streamReader = new StreamReader(fileStream, Encoding.UTF8, true, BufferSize))
            {
                String line;

                while ((line = streamReader.ReadLine()) != null)
                {
                    var record = line.Split("|");
                    dictionary.Add(record[0], record[1]);
                }
            }            

            var queryString = new StringBuilder();
            queryString.AppendLine("SELECT");
            var count = 1;

            foreach(var pair in dictionary)
            {
                bool isDate = false;
                bool isNumeric = false;
                string queryLine = null;
                var stringPattern = "^x([0-9]{2})$";
                var numericPattern = "9";
                var formatLength = 0;
                string numberFormat = null;

                if(Regex.IsMatch(pair.Value, stringPattern))
                {
                    isDate = pair.Key.ToUpper().Contains("DATE") & pair.Value.Equals("x(10)");
                    var startIndex = pair.Value.IndexOf("(") + 1;
                    var endIndex = pair.Value.IndexOf(")");
                    formatLength = int.Parse(pair.Value.Substring(startIndex, endIndex - startIndex));
                }
                else if(Regex.IsMatch(pair.Value, numericPattern))
                {
                    isNumeric = true;
                    formatLength = pair.Value.Length;

                    numberFormat = pair.Value.Replace('>', '#');
                    numberFormat = numberFormat.Replace('9', '0');
                    if (numberFormat[0].Equals("-"))
                    {
                        numberFormat = numberFormat.Replace("-", string.Empty);
                    }
                }
                else
                {
                    //unknown
                }

                queryLine += $"RIGHT(REPLICATE(' ', {formatLength}) + ";

                if (isNumeric)
                {
                    queryLine += $"FORMAT(ISNULL(";
                }

                if (isDate)
                {
                    queryLine += $"FORMAT(";
                }    

                queryLine += $"DATA{count} + ";

                if (isNumeric)
                {
                    queryLine += $", 0), '{numberFormat}'";
                }

                if (isDate)
                {
                    queryLine += ", 'MM/dd/yyyy')";
                }

                queryLine += $"), {formatLength}) ";

                if (pair.Key == dictionary.Last().Key)
                {
                    queryLine += $" as [{pair.Key}]";
                }
                else
                {
                    queryLine += $" as [{pair.Key}],";
                }

                queryString.AppendLine(queryLine);
                count++;
            }

            queryString.AppendLine("FROM");
        }
    }
}
