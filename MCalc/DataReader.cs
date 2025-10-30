using ClosedXML;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Spreadsheet;
using MCalc;
using System.IO;
using System;


namespace MCalc
{
    internal  class DataReader
    {
        readonly  string  _filePath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..\MCalc\Data\AllIn.xlsx"));
        public List<double> Read() 
        {

            if (!File.Exists(_filePath))
            {
                Console.WriteLine($"File not found: {_filePath}");
                return new List<double>();
            }

            var prices = new List<double>();
            using (var wb = new XLWorkbook(_filePath))
            {
                var ws = wb.Worksheet(1);

                int lastRow = ws.LastRowUsed().RowNumber();
                int lastCol = ws.LastColumnUsed().ColumnNumber();

                for (int r = 1; r <= lastRow; r++)
                {
                    for (int c = 2; c <= lastCol; c++)
                    {
                        var value = ws.Cell(r, c).GetValue<string>().Trim();
                        if (double.TryParse(value.Replace(',', '.'), System.Globalization.NumberStyles.Any,
                                            System.Globalization.CultureInfo.InvariantCulture, out double price))
                        {
                            prices.Add(price);
                        }
                    }
                }
            }
            return prices;
        }

        public DataReader() { }
    }
}


