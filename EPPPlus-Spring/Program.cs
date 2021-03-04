using System;

namespace EPPPlus_Spring
{
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;
    using EPPPlus_Spring.Models;
    using OfficeOpenXml;
    using OfficeOpenXml.FormulaParsing;

    internal static class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var results = CalculateSpring();
            Console.WriteLine();
        }

        static (Result leftResult, Result rightResult) CalculateSpring()
        {
            try
            {
                var filePath = new FileInfo(Environment.GetEnvironmentVariable("SPRING_CALCULS_EXCEL_FILE_NAME") ?? string.Empty);

                using var package = new ExcelPackage(filePath);
                var workbook = package.Workbook;

                var inputSheet = workbook.Worksheets["Input"];
                var dataSheet = workbook.Worksheets["Data"];
                var calculationSheet = workbook.Worksheets["Calculation"];

                var inputDoorCells = inputSheet.Cells[12, 1, 12, 6];
                /*
                * Type of door / Calculation / Cycles / Cable Drum / Spring DInside / Number of springs
                * Type of door: Data!Z15
                * Calculation: Data!V8
                * Cycles: Data!N12
                * Cable Drum: Data!AH22
                * Sprint DInside: Data!N22
                * Number of springs: Calculation!B41
                */
                var doorType = dataSheet.Cells["Z15"];
                var calculation = dataSheet.Cells["V8"];
                var cycles = dataSheet.Cells["N12"];
                var cableDrum = dataSheet.Cells["AH22"];
                var springDInside = dataSheet.Cells["N22"];
                var springNumber = calculationSheet.Cells["B41"];

                doorType.Value = 3;
                
                workbook.Calculate(new ExcelCalculationOption {AllowCircularReferences = true});

                return GetResult(inputSheet);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }

        private static (Result leftResult, Result rightResult) GetResult(ExcelWorksheet sheet)
        {
            var resultCells = sheet.Cells[18, 1, 18, 13];
            Console.WriteLine();

            return (null, null);
        }
    }
}