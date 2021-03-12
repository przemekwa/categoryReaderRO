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
        const string pathToFile = @"C:\Users\walkowskip\Downloads\Category tree v1.7.xlsx";
        const string scriptFileName = "categoryReaderRO_script.sql";

        static void Main(string[] args)
        {
            Console.WriteLine("categoryReaderRO v1.0");

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = ReadFile();
            CreateScript(model);

        }

        private static void CreateScript(HashSet<Row> model)
        {
            using var fileStream = new StreamWriter(scriptFileName, false);

            foreach (var item in model)
            {
                foreach (var item2 in item.WUGR)
                {
                    var sb = new StringBuilder();
                    sb.Append("INSERT INTO [SelgrosMainDB_NEW].dbo.[WebsiteCategories] (id_category,wgr,wugr,category) VALUES ");
                    sb.Append("(");
                    sb.Append($"''{Guid.NewGuid()}'', ");
                    sb.Append($"''{item.WGR}'', ");
                    sb.Append($"''{item2}'', ");
                    sb.Append($"N''{item.Category}''");
                    sb.Append(");");

                    fileStream.WriteLine(sb.ToString());
                }
            }
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
                    row.WUGR = new HashSet<int>();

                    if (reader[1] == null)
                    {
                        continue;
                    }

                    row.WGR = reader[1].ToString();
                    row.Category = reader[3].ToString();

                    if (rowList.Contains(row))
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

            foreach (var wugr in reader[2].ToString().Split(',', StringSplitOptions.RemoveEmptyEntries))
            {
                if (int.TryParse(wugr.Trim(), out int result))
                {
                    actualRow.WUGR.Add(result);
                }
                else
                {
                    Console.WriteLine($"wugr {wugr} not well formatted. WGR: {actualRow.WGR}");
                }

            }

            return actualRow;
        }
    }
}
