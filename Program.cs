using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace categoryReaderRO
{
    class Program
    {
        const string pathToFile = @"C:\Users\walkowskip\Downloads\test1.xlsx";
        const string scriptFileName = "categoryReaderRO_script.sql";

        static void Main(string[] args)
        {
            Console.WriteLine("categoryReaderRO v1.0");

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = ReadFile();

        }

        private static HashSet<Row> ReadFile()
        {
            using var stream = File.Open(pathToFile, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            var rowList = new HashSet<Row>();

            do
            {
                reader.Read();

                while (reader.Read())
                {
                    var row = new Row();
                    row.Category = reader[3].ToString();

                    if (rowList.Contains(row,))
                    {
                        rowList.TryGetValue(row, out Row actualRow);
                        SetValue(actualRow, reader);
                    }

                    rowList.Add(SetValue(row, reader));

                }
            } while (reader.NextResult());

            return rowList;
        }

        private static Row SetValue(Row actualRow, IExcelDataReader reader)
        {
            if (actualRow.Id == Guid.Empty)
            {
                actualRow.Id = Guid.NewGuid();
            }

            actualRow.WGR = reader[1].ToString();
            actualRow.WUGR = reader[2].ToString().Split(',').ToList();

            return actualRow;
        }
    }
}
