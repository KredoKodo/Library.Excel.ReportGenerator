using ClosedXML.Excel;
using KredoKodo.Library.Excel.ReportGenerator;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace KredoKodo.Library.Excel.Tests
{
    public abstract class ReportGeneratorTestBase : IDisposable
    {
        protected ReportGeneratorTestBase()
        {
            SampleDataTable = GenerateSampleDataTable();
            SampleReportDataTable = GenerateReportDataTable();
        }

        protected DataTable SampleDataTable { get; set; }
        protected DataTable SampleReportDataTable { get; set; }

        public void Dispose()
        {
            SampleDataTable.Dispose();
            SampleReportDataTable.Dispose();
        }

        private DataTable GenerateSampleDataTable()
        {
            var dataTable = new DataTable();

            dataTable.Columns.Add("Quantity", typeof(int));
            dataTable.Columns.Add("Item", typeof(string));
            dataTable.Columns.Add("Person", typeof(string));
            dataTable.Columns.Add("Date", typeof(DateTime));

            dataTable.Rows.Add(25, "Cookie", "David", default(DateTime));
            dataTable.Rows.Add(50, "Rice", "Stephen", default(DateTime));
            dataTable.Rows.Add(10, "Brat", "Joshua", default(DateTime));
            dataTable.Rows.Add(21, "Chip", "Rev", default(DateTime));
            dataTable.Rows.Add(100000, "$100,000 bar", "Ken", DateTime.Parse("03/03/03 12:59 PM"));

            return dataTable;
        }

        private DataTable GenerateReportDataTable()
        {
            var dataTable = new DataTable();

            dataTable.Columns.Add("Count", typeof(int));
            dataTable.Columns.Add("Dev", typeof(string));
            dataTable.Columns.Add("Food", typeof(string));
            dataTable.Columns.Add("Purchase Date", typeof(DateTime));

            dataTable.Rows.Add(25, "David", "Cookie", default(DateTime));
            dataTable.Rows.Add(50, "Stephen", "Rice", default(DateTime));
            dataTable.Rows.Add(10, "Joshua", "Brat", default(DateTime));
            dataTable.Rows.Add(21, "Rev", "Chip", default(DateTime));
            dataTable.Rows.Add(100000, "Ken", "$100,000 bar", DateTime.Parse("03/03/03 12:59 PM"));

            return dataTable;
        }

        protected XLWorkbook GenerateExpectedWorkbook(
            DataTable dataTable,
            List<ColumnModel> formatList)
        {
            var workbook = new XLWorkbook();
            var configuredWorksheet = workbook
                .AddWorksheet(dataTable, "test");

            configuredWorksheet.Style.Font.FontSize = 10;
            configuredWorksheet
                .Style
                .Font
                .FontFamilyNumbering = XLFontFamilyNumberingValues.Modern;

            foreach (var customFormat in formatList)
            {
                var columnOrdinal = -1;
                foreach (DataColumn column in dataTable.Columns)
                {
                    if (column.ColumnName == customFormat.ColumnName)
                    {
                        columnOrdinal = column.Ordinal;
                    }
                }

                if (columnOrdinal > -1)
                {
                    if (!string.IsNullOrWhiteSpace(customFormat.ColumnFormat))
                    {
                        configuredWorksheet
                            .Column(columnOrdinal + 1)
                            .Style
                            .NumberFormat
                            .Format = customFormat.ColumnFormat;
                    }

                    if (customFormat.IsColumnCentered)
                    {
                        configuredWorksheet
                            .Column(columnOrdinal + 1)
                            .Style
                            .Alignment
                            .Horizontal = XLAlignmentHorizontalValues.Center;
                    }

                    if (!string.IsNullOrWhiteSpace(customFormat.ColumnBackgroundColor))
                    {
                        configuredWorksheet
                            .Column(columnOrdinal + 1)
                            .Style
                            .Fill
                            .BackgroundColor = XLColor.FromHtml(customFormat.ColumnBackgroundColor);

                        // Remove any column formatting on the header row, 
                        // so the header row is the same for all columns
                        configuredWorksheet.Cell(1, columnOrdinal + 1)
                            .Clear(XLClearOptions.AllFormats);
                    }
                }
            }

            return workbook;
        }

        protected DataTable SampleQuantityDataTable()
        {
            var dataTable = new DataTable();

            dataTable.Columns.Add("Count", typeof(decimal));

            dataTable.Rows.Add(25);
            dataTable.Rows.Add(50);
            dataTable.Rows.Add(10);
            dataTable.Rows.Add(21);
            dataTable.Rows.Add(100000);
            dataTable.Rows.Add(35.2);

            return dataTable;
        }

        protected DataTable SampleDatesDataTable()
        {
            var dataTable = new DataTable();

            dataTable.Columns.Add("Service Date", typeof(DateTime));

            dataTable.Rows.Add(default(DateTime));
            dataTable.Rows.Add(DateTime.Parse("03/03/2003"));
            dataTable.Rows.Add(DateTime.Parse("09/05/2010 09:30 PM"));

            return dataTable;
        }

        protected MemoryStream CreateMemoryStream(XLWorkbook workbook)
        {
            var memStream = new MemoryStream();
            workbook.SaveAs(memStream);

            return memStream;
        }
    }
}