using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using OfficeOpenXml;

namespace excel_parser
{
    class Program
    {
        static void Main(string[] args)
        {
            var filePath = @"./data/data.xlsx";
            FileInfo file = new FileInfo(filePath);

            var exportList = new List<ExportClass>();

            using (var package = new ExcelPackage(file))
            {
                var worksheet = package.Workbook.Worksheets[1];
                var rowCount = worksheet.Dimension.Rows;
                var ColCount = worksheet.Dimension.Columns;

                var rawText = string.Empty;
                for (int row = 1; row <= rowCount; row++)
                {
                    for (var col = 1; col <= ColCount; col++)
                    {
                        rawText += worksheet.Cells[row, col].Value.ToString(); //+ "\t";
                    }
                    // rawText += "\r\n";
                    exportList.Add(new ExportClass(){Name = rawText});
                }
                // Console.WriteLine(rawText);
            }

            var export = JsonConvert.SerializeObject(exportList);
            Console.WriteLine(export);

        }
    }
}
