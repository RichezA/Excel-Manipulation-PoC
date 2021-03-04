using System;

namespace EPPPlus_Spring
{
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Reflection.Metadata;
    using System.Threading.Tasks;
    using EPPPlus_Spring.Models;
    using OfficeOpenXml;
    using OfficeOpenXml.FormulaParsing;

    internal static class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var filePath = new FileInfo(Environment.GetEnvironmentVariable("SPRING_CALCULS_EXCEL_FILE_NAME") ?? string.Empty);
            using var package = new ExcelPackage(filePath);
            Stopwatch watcher = new Stopwatch();

            watcher.Start();
            Console.WriteLine(watcher.Elapsed);
            var results = package.Workbook.CalculateSpring(2, 2);
            Console.WriteLine(watcher.Elapsed);
            watcher.Stop();
        }

        static (Result leftResult, Result rightResult) CalculateSpring(this ExcelWorkbook workbook, int doorValue, int numberOfSprings)
        {
            try
            {

                // Fetch used sheets
                var inputSheet = workbook.Worksheets["Input"];
                var dataSheet = workbook.Worksheets["Data"];
                var calculationSheet = workbook.Worksheets["Calculation"];

                // Fetch Inputs
                var doorWidth = inputSheet.Cells["B12"];
                var doorHeight = inputSheet.Cells["C12"];
                var doorLift = inputSheet.Cells["D12"];
                var doorPitch = inputSheet.Cells["E12"];
                var doorWeight = inputSheet.Cells["F12"];
                var doorBs = inputSheet.Cells["G12"];
                
                // preference inputs
                var doorType = dataSheet.Cells["Z15"];
                var calculation = dataSheet.Cells["V8"];
                var cycles = dataSheet.Cells["N12"];
                var cableDrum = dataSheet.Cells["AH22"];
                var springDInside = dataSheet.Cells["N22"];
                var springNumber = calculationSheet.Cells["B41"];

                // Modify input values.
                doorType.Value = doorValue;
                springNumber.Value = numberOfSprings;

                // Fetch results
                var results = GetResult(inputSheet);

                // Call same method with another preference (doorType, calculation, ..., springNumber) if no solutions are available.
                return results.leftResult.Spring.Contains("No solution") ||
                       results.leftResult.Spring.Contains("not a stock item") ?
                    workbook.CalculateSpring(2, 2) :
                    results;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }
        
        private static void CalculateCells(ExcelRangeBase excelRange)
        {
            excelRange.Calculate(new ExcelCalculationOption { AllowCircularReferences = true});
        }

        private static (Result leftResult, Result rightResult) GetResult(ExcelWorksheet sheet)
        {
            // Fetch cells
            var resultCells = sheet.Cells[18, 1, 19, 13];
            
            // Calculate result cells
            CalculateCells(resultCells);

            // Fetch values.
            var resultValues = GetCellValues(resultCells, removeNull: true);
            var leftResultValues = resultValues.Take(resultValues.Count / 2);
            var rightResultValues = resultValues.Skip(resultValues.Count / 2);
            
            // Return result.
            return (GetResult(leftResultValues), GetResult(rightResultValues));
        }


        private static Result GetResult(IEnumerable<object?> propertiesToMap)
        {
            var props = propertiesToMap.ToArray();
            if (typeof(Result).GetProperties().Length < props.Length)
            {
                return null;
            }

            if (double.TryParse(props[0]?.ToString(), out var number) &&
                props[1]?.ToString() is { } coilDir &&
                double.TryParse(props[2]?.ToString(), out var id) &&
                props[3]?.ToString() is { } spring &&
                props[4] is { } length &&
                double.TryParse(props[5]?.ToString(), out var turns) &&
                double.TryParse(props[6]?.ToString(), out var weight) &&
                double.TryParse(props[7]?.ToString(), out var space) &&
                double.TryParse(props[8]?.ToString(), out var freeSpace))
            {
                return new Result
                {
                    Number = number,
                    CoilDir = coilDir,
                    Id = id,
                    Spring = spring,
                    Length = length,
                    Turns = turns,
                    Weight = weight,
                    Space = space,
                    FreeSpace = freeSpace
                };
            }

            return null;
        }
        
        private static List<object?> GetCellValues(ExcelRange range, bool removeNull = false)
            => range.Select(c => c.GetValue<object?>()).Where(c => c != null).ToList();

    }
}