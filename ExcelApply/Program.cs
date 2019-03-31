using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace ExcelApply
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 5)
            {
                Console.WriteLine("5 arguments expected in the following order:");
                Console.WriteLine("*input filename* *output filename* *worksheet name* *cell for formula to apply* *range to apply formula to*");
                Console.WriteLine("An example: in.xlsx out.csv averages C3 D4:Z999");
                Console.WriteLine();
                Console.WriteLine("There's no error checking whatsoever aside from making sure there're 5 arguments...");
                return;
            }

            var excelFile = args[0];
            var outFile = args[1];
            var worksheet = args[2];
            var source = args[3];
            var targetRange = args[4];

            string FormatValueForOutput(object o)
            {
                switch (o)
                {
                    case float f:
                        return f.ToString("F2");
                    case double d:
                        return d.ToString("F2");
                    case DateTime dt:
                        return dt.ToString("yyyy.MM.");
                    case string s:
                        return s;
                    default:
                        return o.ToString();
                }
            }

            Console.WriteLine("Loading excel file...");

            using (var ep = new ExcelPackage(new FileInfo(excelFile)))
            {
                Console.WriteLine("File loaded.");

                var formula = ep.Workbook.Worksheets[worksheet].Cells[source].FormulaR1C1;
                
                var range = ep.Workbook.Worksheets[worksheet].Cells[targetRange];

                for (var i = range.Start.Row; i <= range.End.Row; i++)
                {
                    for (var j = range.Start.Column; j <= range.End.Column; j++)
                    {
                        ep.Workbook.Worksheets[worksheet].Cells[i, j].FormulaR1C1 = formula;
                    }

                    Console.WriteLine($"Applying formula to line {i}...");
                }

                if (File.Exists(outFile)) File.Delete(outFile);

                using (var f = File.OpenWrite(outFile))
                {
                    using (var sw = new StreamWriter(f))
                    {
                        for (var i = 1; i <= range.End.Row; i++)
                        {
                            Console.WriteLine($"Working on line {i}...");
                            var rowCells = ep.Workbook.Worksheets[worksheet].Cells[$"A{i}:ATR{i}"];
                            rowCells.Calculate();
                            foreach (var rowCell in rowCells.SkipLast(1))
                            {
                                sw.Write(FormatValueForOutput(rowCell.Value));
                                sw.Write(";");
                            }
                            sw.WriteLine(FormatValueForOutput(rowCells.Last().Value));
                        }
                    }
                }
            }
        }
    }
}
