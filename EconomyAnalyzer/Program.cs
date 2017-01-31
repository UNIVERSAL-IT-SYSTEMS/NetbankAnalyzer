using EconomyAnalyzer.Entitites;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EconomyAnalyzer
{
    class Program
    {
        static void Main(string[] args)
        {
            // Read file (csv parser)
            var path = @"c:\temp\";

            var categories = File.ReadAllLines(path + "categories.txt", Encoding.Default).ToArray();
            var i = 0;
            foreach (var category in categories)
            {
                Console.WriteLine($"{i++}: {category}");
            }
            Console.WriteLine("---------------------------------------");

            var links = File.ReadAllLines(path + "links.txt", Encoding.Default).Select(s => {
                var split = s.Split('|');
                return new { Key = split[0], Value = split[1] };
            })
            .ToDictionary(s => s.Key,s => s.Value);

            var lines = File.ReadAllLines(path + "poster.csv", Encoding.Default);

            // the first two lines are empty and the header
            var rows = lines.Skip(2).Select(s => new Row(s)).Where(s => s != null).ToList();

            UpdateRows(rows, categories, links);

            // map lines to category
            while (rows.Where(s => string.IsNullOrEmpty(s.Category)).Any())
            {
                var row = rows.Where(s => string.IsNullOrEmpty(s.Category)).First();

                if (links.ContainsKey(row.SanitizedComment))
                {
                    row.Category = links[row.SanitizedComment];
                    continue;
                }
                var count = rows.Count(s => string.IsNullOrEmpty(s.Category));
                // hacks to clear the line
                Console.Write($"\r({count}) Hvad er det: {row.SanitizedComment}                                     ");
                var result = Console.ReadKey();
                var index = int.Parse(result.KeyChar.ToString());

                row.Category = categories[index];
                links[row.SanitizedComment] = row.Category;

                File.AppendAllLines(path + "links.txt", new List<string> { $"{row.SanitizedComment}|{row.Category}" });
                UpdateRows(rows, categories, links);
            }

            rows = rows.OrderBy(s => s.Date).ThenByDescending(s => s.Balance).ToList();

            // output to xsls file
            var package = new ExcelPackage();
            ExcelWorksheet currentSheet = null;
            int currentRow = 2;
            var sheetName = "";


            foreach (var item in rows)
            {
                var newSheetName = $"{item.Date.ToString("MMMMM")} {item.Date.Year}";


                if (sheetName != newSheetName)
                {
                    if(currentSheet != null)
                    {
                        for (int j = 0; j <= categories.Length; j++)
                        {
                            var index = j + 3;

                            // Capital A starts at char value 65, so we start one before that
                            var letter = Char.ConvertFromUtf32(index + 64);
                            currentSheet.Cells[currentRow, index].Formula = $"=SUM({letter}2:{letter}{currentRow-1})";
                        }

                        // Make the totals bold
                        using (var range = currentSheet.Cells[currentRow, 1, currentRow, 3 + categories.Length])
                        {
                            range.Style.Font.Bold = true;
                        }


                        // Select all the cells to adjust cell width
                        using (var range = currentSheet.Cells["A1:L500"])
                        {
                            range.AutoFitColumns();
                        }
                    }

                    sheetName = newSheetName;
                    currentSheet = package.Workbook.Worksheets.Add(sheetName);

                    // headers
                    currentSheet.Cells[1, 1].Value = "Dato";
                    currentSheet.Cells[1, 2].Value = "Text";
                    currentSheet.Cells[1, 3].Value = "Ændring";

                    var start = 4;
                    foreach (var category in categories)
                    {
                        currentSheet.Cells[1, start++].Value = category;
                    }

                    currentRow = 2;
                }

                currentSheet.Cells[currentRow, 1].Value = item.Date;
                currentSheet.Cells[currentRow, 1].Style.Numberformat.Format = "yyyy-mm-dd";
                currentSheet.Cells[currentRow, 2].Value = item.SanitizedComment;
                currentSheet.Cells[currentRow, 3].Value = item.Amount;

                var categoryIndex = categories.ToList().IndexOf(item.Category) + 1;
                currentSheet.Cells[currentRow, 3 + categoryIndex].Value = Math.Abs(item.Amount);

                currentRow++;
            }

            File.WriteAllBytes(path + $"result {DateTime.Now.ToString("yyyy-MM-dd HHmm")}.xlsx", package.GetAsByteArray());

            //// Add headers
            //var content = new List<string>();
            //content.Add("Dato,Tekst,Beløb,Saldo," + string.Join(",", categories));

            //content.AddRange(rows.Select(s => s.ToString(categories)));


            //var newFile = string.Join(Environment.NewLine, content);
            //File.WriteAllText(path + "posterout.csv",  newFile, Encoding.Default);
        }

        private static void UpdateRows(List<Row> rows, string[] categories, Dictionary<string, string> links)
        {
            foreach (var row in rows)
            {
                if (!string.IsNullOrEmpty(row.Category))
                {
                    continue;
                }

                if (links.ContainsKey(row.SanitizedComment))
                {
                    row.Category = links[row.SanitizedComment];
                }
            }
        }
    }
}
