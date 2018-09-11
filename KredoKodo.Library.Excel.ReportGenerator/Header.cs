using ClosedXML.Excel;
using System;
using System.Linq;

namespace KredoKodo.Library.Excel.ReportGenerator
{
    /// <summary>
    /// Provides methods for adding a header to all Excel worksheets
    /// </summary>
    public static class Header
    {
        private const string PROPERTY_KEY = "KredoKodo.Library.Excel.ReportGenerator";

        /// <summary>
        /// Adds a header row, with the specified string, to all 
        /// worksheets within a workbook
        /// </summary>
        /// <param name="workbook">The ClosedXML Workbook to operate on</param>
        /// <param name="header">The string to use</param>
        public static void Mark(XLWorkbook workbook, string header)
        {
            CheckWorkbookHeaderExists(workbook);
            CheckWorksheetsExist(workbook);

            foreach (IXLWorksheet worksheet in workbook.Worksheets)
            {
                worksheet.Row(1).InsertRowsAbove(1);

                var lastColumnUsed = worksheet
                            .ColumnsUsed()
                            .Last()
                            .ColumnLetter() + "1";
                var markingCell = worksheet.Cell(1, 1);

                markingCell.Value = header;
                markingCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                markingCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                markingCell.Style.Font.Bold = true;
                worksheet.Row(1).Height = 40;
                worksheet.Range($"A1:{lastColumnUsed}").Merge();
            }

            AddHeaderProperty(workbook);
        }

        private static void CheckWorkbookHeaderExists(XLWorkbook workbook)
        {
            // If the property exists at all then throw; the value doesn't 
            // matter and is for reference.

            var propertyExists = workbook
                .CustomProperties
                .FirstOrDefault(p => p.Name == PROPERTY_KEY);

            if (propertyExists != null)
            {
                throw new InvalidOperationException(
                    "The workbook has already been operated on.");
            }
        }

        private static void CheckWorksheetsExist(XLWorkbook workbook)
        {
            if (!workbook.Worksheets.Any())
            {
                throw new InvalidOperationException(
                    "The workbook requires at least 1 worksheet.");
            }
        }

        private static void AddHeaderProperty(XLWorkbook workbook)
        {
            workbook.CustomProperties.Add(PROPERTY_KEY, DateTime.Now);
        }
    }
}
