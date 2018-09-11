using ClosedXML.Excel;
using System;
using System.Data;
using System.IO;

namespace KredoKodo.Library.Excel.ReportGenerator
{
    /// <summary>
    /// Provides methods for converting DataTables into Excel worksheets for
    /// the purposes of easily creating basic Excel workbook reports.
    /// </summary>
    /// 
    /// <remarks>
    /// Example usage:
    /// 
    /// var reportGenerator = new ReportGenerator();
    ///
    /// reportGenerator
    ///  .AddWorksheet(myDataTable, myWorksheetName)
    ///      .ConfigureColumn("dt_col_1", "First Name", ...options)
    ///      .ConfigureColumn("dt_col_2", "Last Name", ...options)
    ///      .FinalizeWorksheet()
    ///  .AddWorksheet(myOtherTable, mySecondSheet)
    ///      .ConfigureColumn("dt_col_1", "Company Name", ...options)
    ///      .ConfigureColumn("dt_col_2", "Street Address", ...options)
    ///      .FinalizeWorksheet()
    ///  .GenerateWorkbook("optional header") or
    ///  .GenerateWorkbookBytes("optional header");
    /// 
    /// </remarks>
    public class ReportGenerator : IDisposable
    {
        /// <summary>
        /// Initializes a new instance of the 
        /// KredoKodo.Library.Excel.ReportGenerator class.
        /// </summary>
        public ReportGenerator()
        {
            WorkbookReport = new XLWorkbook();
        }

        /// <summary>
        /// Gets and sets the ClosedXml workbook in which generated worksheets
        /// will be added to and finally returned.
        /// </summary>
        public XLWorkbook WorkbookReport { get; protected set; }

        /// <summary>
        /// Creates a worksheet in which the DataTable data will be appended.
        /// </summary>
        /// <param name="dataTable">The DataTable that holds the worksheet data.</param>
        /// <param name="worksheetName">The name of the worksheet.</param>
        /// <returns></returns>
        public WorksheetModel AddWorksheet(
            DataTable dataTable,
            string worksheetName)
        {
            if (string.IsNullOrWhiteSpace(worksheetName))
            {
                throw new ArgumentException(
                    "You must provide a worksheet name.");
            }

            if (dataTable.Columns.Count < 1)
            {
                throw new InvalidOperationException(
                    "You must provide a DataTable with at least one column.");
            }

            return new WorksheetModel(dataTable, worksheetName, this);
        }

        public void Dispose()
        {
            WorkbookReport.Dispose();
        }

        /// <summary>
        /// A workbook consisting of all configured worksheets.
        /// </summary>
        /// <returns>A ClosedXml workbook.</returns>
        public XLWorkbook GenerateWorkbook(string banner = "")
        {
            AddBanner(banner);

            return WorkbookReport;
        }

        /// <summary>
        /// A workbook consisting of all configured worksheets.
        /// </summary>
        /// <returns>A ClosedXml workbook as a byte array.</returns>
        public byte[] GenerateWorkbookBytes(string banner = "")
        {
            byte[] workbookBytes = new byte[0];

            AddBanner(banner);

            using (var memoryStream = new MemoryStream())
            {
                WorkbookReport.SaveAs(memoryStream);
                workbookBytes = memoryStream.ToArray();
            }

            return workbookBytes;
        }

        private void AddBanner(string banner)
        {
            if (!string.IsNullOrWhiteSpace(banner))
            {
                Header.Mark(WorkbookReport, banner);
            }
        }
    }
}