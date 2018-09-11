using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace KredoKodo.Library.Excel.ReportGenerator
{
    /// <summary>
    /// Represents one worksheet within an Excel workbook
    /// </summary>
    public class WorksheetModel
    {
        private DataTable dataTable;
        private List<string> outputColumnNames;
        private readonly string worksheetName;
        private List<ColumnModel> columnModels;
        private ReportGenerator parent;
        private const int DESIGNATION_AND_HEADER_ROWS = 2;

        /// <summary>
        /// Initializes a new instance of the 
        /// KredoKodo.Library.Excel.ReportGenerator.WorksheetModel
        /// class with the specified datatable and worksheet name.
        /// </summary>
        /// <param name="dataTable">The datatable to operated on.</param>
        /// <param name="worksheetName">The name of the worksheet.</param>
        /// <param name="parentReportGenerator">The reference to the Report 
        ///         Generator class tied to the CustomWorksheetInfo.</param>
        public WorksheetModel(
            DataTable dataTable,
            string worksheetName,
            ReportGenerator parentReportGenerator)
        {
            columnModels = new List<ColumnModel>();
            outputColumnNames = new List<string>();
            this.dataTable = dataTable;
            this.worksheetName = worksheetName;
            parent = parentReportGenerator;
        }

        /// <summary>
        /// Gets the table of data for insert into the worksheet and
        /// filtering out unused.
        /// </summary>
        private DataTable WorksheetDataTable
        {
            get
            {
                var view = new DataView(dataTable);
                return view.ToTable(false, outputColumnNames.ToArray());
            }
        }

        /// <summary>
        /// Used to change the column names and order for a given dataset when
        /// sending that dataset to generate an excel spreadsheet.
        /// </summary>
        /// <param name="dataTableName">The column name from the dataset.</param>
        /// <param name="reportColumnName">The column name you want in the spreadsheet.</param>
        /// <param name="specialFormat">Use empty string or choose a CustomFormat.</param>
        /// <param name="isCentered">True centers the entire column.</param>
        /// <param name="backgroundColor">Use empty string or choose a CustomExcelColor.</param>
        /// <returns></returns>
        public WorksheetModel ConfigureColumn(
            string dataTableName,
            string reportColumnName,
            string specialFormat = "",
            bool isCentered = false,
            string backgroundColor = "")
        {
            if (!dataTable.Columns.Contains(dataTableName))
            {
                throw new ArgumentException(
                    "You must provide a matching DataColumn name as the first parameter.");
            }

            if (string.IsNullOrWhiteSpace(reportColumnName))
            {
                throw new ArgumentException(
                    "You must provide a valid Excel column name as the second parameter.");
            }

            bool hasSpecialFormat = !string.IsNullOrWhiteSpace(specialFormat);
            bool hasCustomColor = !string.IsNullOrWhiteSpace(backgroundColor);

            if (hasSpecialFormat || isCentered || hasCustomColor)
            {
                var customColumnFormat = new ColumnModel
                {
                    ColumnName = reportColumnName
                };

                if (hasSpecialFormat)
                    customColumnFormat.ColumnFormat = specialFormat;

                if (isCentered)
                    customColumnFormat.IsColumnCentered = true;

                if (hasCustomColor)
                    customColumnFormat.ColumnBackgroundColor = backgroundColor;

                columnModels.Add(customColumnFormat);
            }

            dataTable.Columns[dataTableName].ColumnName = reportColumnName;
            outputColumnNames.Add(reportColumnName);
            dataTable.AcceptChanges();

            return this;
        }

        /// <summary>
        /// Formats the worksheet in accordance with all of the applied
        /// column configurations before adding the worksheet to the workbook.
        /// </summary>
        ///
        /// <returns>The parent Report Generator.</returns>
        public ReportGenerator FinalizeWorksheet(double fontSize = 10)
        {
            if (!outputColumnNames.Any())
            {
                throw new InvalidOperationException(
                    "You must configure at least one column.");
            }

            var configuredWorksheet = parent.WorkbookReport
                .AddWorksheet(WorksheetDataTable, worksheetName);

            configuredWorksheet.Style.Font.FontSize = fontSize;
            configuredWorksheet.Style.Font.FontFamilyNumbering = XLFontFamilyNumberingValues.Modern;

            foreach (var customFormat in columnModels)
            {
                var columnOrdinal = -1;
                foreach (DataColumn column in WorksheetDataTable.Columns)
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

            configuredWorksheet
                .SheetView
                .FreezeRows(DESIGNATION_AND_HEADER_ROWS);

            return parent;
        }
    }
}