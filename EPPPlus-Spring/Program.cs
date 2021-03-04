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
            var results = package.Workbook.CalculateSpring(3, 1);
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

        private static (Result leftResult, Result rightResult) GetResult(ExcelWorksheet sheet)
        {
            // Fetch cells
            var leftResultCells = sheet.GetCells(18, 1, 18, 13);
            var rightResultCells = sheet.GetCells(19, 1, 19, 13);
            
            // Calculate result cells
            CalculateCells(sheet.Cells[18, 1, 19, 13]);

            // Fetch values.
            var leftResultValues = GetCellValues(leftResultCells, removeNull: true);
            var rightResultValues = GetCellValues(rightResultCells, removeNull: true);
            
            // Return result.
            return (GetResult(leftResultValues), GetResult(rightResultValues));
        }

        private static void CalculateCells(ExcelRangeBase excelRange)
        {
            excelRange.Calculate(new ExcelCalculationOption {AllowCircularReferences = true});
        }

        private static ExcelRange GetCells(this ExcelWorksheet sheet, int fromRow, int fromCol, int toRow,
            int toCol)
        {
            return sheet.Cells[fromRow, fromCol, toRow, toCol];
        }

        private static List<object?> GetCellValues(ExcelRange range, bool removeNull = false)
            => range.Select(c => c.GetValue<object?>()).Where(c => c != null).ToList();

        private static Result GetResult(List<object?> propertiesToMap)
        {
            var result = new Result();
            var properties = result.GetType().GetProperties();

            if (properties.Length != propertiesToMap.Count())
            {
                return null;
            }

            foreach (var (prop, index) in propertiesToMap.Select((val, ind) => (val, ind)))
            {
                var property = properties[index];
                property.SetValue(result, prop);
            }

            return result;
        }
    }
}