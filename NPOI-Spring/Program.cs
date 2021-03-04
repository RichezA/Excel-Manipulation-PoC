using System;

namespace NPOI_Spring
{
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.XSSF.UserModel;

    static class Program
    {
        private static readonly string Directory = @"C:\Users\a.richez\Codes\QSet\";
        private static readonly string ExcelPath = $@"{Directory}Copy of AlcomexTor.xls";
        static void Main(string[] args)
        {
            var workbook = GetWorkbook(ExcelPath);
            
            var inputSheet = workbook.GetSheet("Input");

            if (inputSheet == null)
            {
                Console.WriteLine("Could not find input.");
                return;
            }

            // door infos
            var inputDoorCells = inputSheet.GetRow(11).GetCells(6);
            /*
            * Type of door / Calculation / Cycles / Cable Drum / Spring DInside / Number of springs
            * Type of door: Data!Z15
            * Calculation: Data!V8
            * Cycles: Data!N12
            * Cable Drum: Data!AH22
            * Sprint DInside: Data!N22
            * Number of springs: Calculation!B41
            */
            var dataSheet = workbook.GetSheet("Data");
            var data = dataSheet.GetCellsFromReference(new[] {"Z15", "V8", "N12", "AH22", "N22"}).ToList();
            var calculationSheet = workbook.GetSheet("Calculation");
            var springNumber = calculationSheet.GetCellFromReference("B41");
            var results = workbook.GetSheet("Input").GetRows(new[] { 17, 18 }).Select(x => x.GetCells(9));
            
            data.First().SetCellValue((int)Options.TypeOfDoorOperation.MechanicalIndustrial1And1_4Inches);
            // workbook.EvaluateAll();
            
            // Compare results to excel.
            var tempPath = $"{Directory}AlcomexTor-Modified_{Guid.NewGuid()}.xls";
            var fStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write);
            workbook.Write(fStream);
            fStream.Dispose();


            var modified = new FileStream(tempPath, FileMode.Open, FileAccess.ReadWrite);
            var modifiedWorkbook = new HSSFWorkbook(modified);
            
            modifiedWorkbook.EvaluateAll();

            var test = modifiedWorkbook.GetSheet("Input").GetRows(new[] { 17, 18 }).Select(x => x.GetCells(9));

            Console.WriteLine();
            File.Delete(tempPath);
        }

        static void EvaluateAll(this IWorkbook workbook)
        {
            var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
            for (var i = 0; i < workbook.NumberOfSheets; i++)
            {
                var sheet = workbook.GetSheetAt(i);
                foreach (IRow row in sheet)
                {
                    foreach (var cell in row)
                    {
                        if (cell.CellType != CellType.Formula) continue;
                        var cellValue = evaluator.Evaluate(cell);
                        if (cellValue.CellType != CellType.Numeric) continue;
                        var cached = cell.NumericCellValue;
                        var evaluated = cellValue.NumberValue;
                        Console.WriteLine($"Cache: {cached} | Evaluated: {evaluated}");
                    }
                }
            }
            
            XSSFFormulaEvaluator.EvaluateAllFormulaCells(workbook);
        }

        static object GetCellValue(this ICell cell)
        {
            return cell.CellType switch
            {
                CellType.Numeric => cell.NumericCellValue,
                CellType.String => cell.StringCellValue,
                CellType.Boolean => cell.BooleanCellValue,
                CellType.Error => cell.ErrorCellValue,
                (CellType.Formula | CellType.Blank | CellType.Unknown) => cell.ToString(),
                _ => null
            };
        }

        static ICell GetCellFromReference(this ISheet sheet, string reference)
        {
            var cellRef = new CellReference(reference);
            
            var row = sheet.GetRow(cellRef.Row);
            return row.GetCell(cellRef.Col);
        }

        static IEnumerable<ICell> GetCellsFromReference(this ISheet sheet, IEnumerable<string> references)
        {
            // var cellRefs = references.Select(r => new CellReference(r));
            // return (from cellRef in cellRefs let row = sheet.GetRow(cellRef.Row) select row.GetCell(cellRef.Col)).ToList();
            return references.Select(r => GetCellFromReference(sheet, r));
        }
        
        static IEnumerable<IRow> GetRows(this ISheet sheet, IEnumerable<int> ids)
            => ids.Select(id => sheet.GetRow(id));

        static IEnumerable<ICell> GetCells(this IRow row, int count)
            => row.Cells.Where(c => c.CellType != CellType.Blank).Take(count);

        static IWorkbook GetWorkbook(string path)
        {
            try
            {
                using var fs = new FileStream(path, FileMode.Open, FileAccess.ReadWrite);

                return new HSSFWorkbook(fs);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }

        static void EvaluateAllFormulas(this HSSFWorkbook workbook)
        {
            var formulaEvaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
            foreach (var sheet in workbook)
            {
                foreach (IRow row in sheet)
                {
                    foreach (var cell in row)
                    {
                        if (cell.CellType == CellType.Formula)
                        {
                            formulaEvaluator.EvaluateFormulaCell(cell);
                        }
                    }
                }
            }
        }
    }
}