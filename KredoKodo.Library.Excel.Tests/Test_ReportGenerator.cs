using ClosedXML.Excel;
using KredoKodo.Library.Excel.ReportGenerator;
using Shouldly;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Xunit;

namespace KredoKodo.Library.Excel.Tests
{
    [Trait("Type", "ReportGenerator Excel Tests")]
    public class Test_ReportGenerator : ReportGeneratorTestBase
    {
        private readonly ReportGenerator.ReportGenerator subject;

        // TODO: Build and add equality library.
        //private readonly ExcelEquality excelEquality;

        public Test_ReportGenerator() : base()
        {
            subject = new ReportGenerator.ReportGenerator();
            //excelEquality = new ExcelEquality();
        }

        [Fact(DisplayName = "ReportGenerator: a datatable with no columns throws")]
        public void RequiresAtLeastOneColumn()
        {
            var emptyDatatable = new DataTable();

            Should.Throw<InvalidOperationException>(() =>
                subject
                    .AddWorksheet(emptyDatatable, "test")
                        .FinalizeWorksheet()
                    .GenerateWorkbook())
                .Message
                .ShouldBe(
                    "You must provide a DataTable with at least one column.");
        }

        [Fact(DisplayName = "ReportGenerator: ConfigureColumn requires at least one column")]
        public void ConfigureColumnRequiresAtLeastOneColumn()
        {
            Should.Throw<InvalidOperationException>(() =>
                subject
                    .AddWorksheet(SampleDataTable, "test")
                        .FinalizeWorksheet())
            .Message
            .ShouldBe("You must configure at least one column.");
        }

        [Fact(DisplayName = "ReportGenerator: requires a worksheet name")]
        public void RequireWorksheetName()
        {
            Should.Throw<ArgumentException>(() =>
               subject.AddWorksheet(SampleDataTable, ""))
            .Message
            .ShouldBe("You must provide a worksheet name.");
        }

        [Fact(DisplayName = "ReportGenerator: GenerateWorkbook returns XLWorkbook")]
        public void GenerateWorkbookReturnsXLWorkbook()
        {
            var testWorkbook = subject
                .AddWorksheet(SampleQuantityDataTable(), "test")
                    .ConfigureColumn("Count", "Count")
                    .FinalizeWorksheet()
                .GenerateWorkbook();

            testWorkbook.ShouldBeOfType<XLWorkbook>();
        }

        [Fact(DisplayName = "ReportGenerator: GenerateWorkbook returns XLWorkbook bytes")]
        public void GenerateWorkbookReturnsXLWorkbookBytes()
        {
            var testWorkbook = subject
                .AddWorksheet(SampleQuantityDataTable(), "test")
                    .ConfigureColumn("Count", "Count")
                    .FinalizeWorksheet()
                .GenerateWorkbookBytes();

            testWorkbook.ShouldBeOfType<byte[]>();
        }

        //[Fact(DisplayName = "ReportGenerator: GenerateWorkbook returns quantity formatted XLWorkbook")]
        //public void GenerateWorkbookReturnsQuantityFormattedXLWorkbook()
        //{
        //    var formatList = new List<ColumnModel>()
        //    {
        //        new ColumnModel
        //        {
        //            ColumnFormat = CustomFormat.Quantity,
        //            ColumnName = "Count"
        //        }
        //    };

        //    var testWorkbook = subject
        //        .AddWorksheet(SampleQuantityDataTable(), "test")
        //            .ConfigureColumn("Count", "Count", CustomFormat.Quantity)
        //            .FinalizeWorksheet()
        //        .GenerateWorkbook();

        //    var streams = new Stream[]
        //    {
        //        CreateMemoryStream(testWorkbook),
        //        CreateMemoryStream(GenerateExpectedWorkbook(
        //            SampleQuantityDataTable(),
        //            formatList))
        //    };

        //    excelEquality
        //        .StreamsAreEqual(streams)
        //        .AreEqual
        //        .ShouldBeTrue();
        //}

        //[Fact(DisplayName = "ReportGenerator: GenerateWorkbook returns currency 0 decimal formatted XLWorkbook")]
        //public void GenerateWorkbookReturns0DecimalFormattedXLWorkbook()
        //{
        //    var formatList = new List<ColumnModel>()
        //    {
        //        new ColumnModel
        //        {
        //            ColumnFormat = CustomFormat.Currency_0_Decimals,
        //            ColumnName = "Count"
        //        }
        //    };

        //    var testWorkbook = subject
        //        .AddWorksheet(SampleQuantityDataTable(), "test")
        //            .ConfigureColumn("Count", "Count", CustomFormat.Currency_0_Decimals)
        //            .FinalizeWorksheet()
        //        .GenerateWorkbook();

        //    var streams = new Stream[]
        //    {
        //        CreateMemoryStream(testWorkbook),
        //        CreateMemoryStream(GenerateExpectedWorkbook(
        //            SampleQuantityDataTable(),
        //            formatList))
        //    };

        //    excelEquality
        //        .StreamsAreEqual(streams)
        //        .AreEqual
        //        .ShouldBeTrue();
        //}

        //[Fact(DisplayName = "ReportGenerator: GenerateWorkbook returns currency 2 decimal formatted XLWorkbook")]
        //public void GenerateWorkbookReturns2DecimalFormattedXLWorkbook()
        //{
        //    var formatList = new List<ColumnModel>()
        //    {
        //        new ColumnModel
        //        {
        //            ColumnFormat = CustomFormat.Currency_2_Decimals,
        //            ColumnName = "Count"
        //        }
        //    };

        //    var testWorkbook = subject
        //        .AddWorksheet(SampleQuantityDataTable(), "test")
        //            .ConfigureColumn("Count", "Count", CustomFormat.Currency_2_Decimals)
        //            .FinalizeWorksheet()
        //        .GenerateWorkbook();

        //    var streams = new Stream[]
        //    {
        //        CreateMemoryStream(testWorkbook),
        //        CreateMemoryStream(GenerateExpectedWorkbook(
        //            SampleQuantityDataTable(),
        //            formatList))
        //    };

        //    excelEquality
        //        .StreamsAreEqual(streams)
        //        .AreEqual
        //        .ShouldBeTrue();
        //}

        //[Fact(DisplayName = "ReportGenerator: GenerateWorkbook returns currency 4 decimal formatted XLWorkbook")]
        //public void GenerateWorkbookReturns4DecimalFormattedXLWorkbook()
        //{
        //    var formatList = new List<ColumnModel>()
        //    {
        //        new ColumnModel
        //        {
        //            ColumnFormat = CustomFormat.Currency_4_Decimals,
        //            ColumnName = "Count"
        //        }
        //    };

        //    var testWorkbook = subject
        //        .AddWorksheet(SampleQuantityDataTable(), "test")
        //            .ConfigureColumn("Count", "Count", CustomFormat.Currency_4_Decimals)
        //            .FinalizeWorksheet()
        //        .GenerateWorkbook();

        //    var streams = new Stream[]
        //    {
        //        CreateMemoryStream(testWorkbook),
        //        CreateMemoryStream(GenerateExpectedWorkbook(
        //            SampleQuantityDataTable(),
        //            formatList))
        //    };

        //    excelEquality
        //        .StreamsAreEqual(streams)
        //        .AreEqual
        //        .ShouldBeTrue();
        //}

        //[Fact(DisplayName = "ReportGenerator: GenerateWorkbook returns shortdate formatted XLWorkbook")]
        //public void GenerateWorkbookReturnsShortDateFormattedXLWorkbook()
        //{
        //    var formatList = new List<ColumnModel>()
        //    {
        //        new ColumnModel
        //        {
        //            ColumnFormat = CustomFormat.ShortDate,
        //            ColumnName = "Service Date"
        //        }
        //    };

        //    var testWorkbook = subject
        //        .AddWorksheet(SampleDatesDataTable(), "test")
        //            .ConfigureColumn("Service Date", "Service Date", CustomFormat.ShortDate)
        //            .FinalizeWorksheet()
        //        .GenerateWorkbook();

        //    var streams = new Stream[]
        //    {
        //        CreateMemoryStream(testWorkbook),
        //        CreateMemoryStream(GenerateExpectedWorkbook(
        //            SampleDatesDataTable(),
        //            formatList))
        //    };

        //    excelEquality
        //        .StreamsAreEqual(streams)
        //        .AreEqual
        //        .ShouldBeTrue();
        //}

        //[Fact(DisplayName = "ReportGenerator: GenerateWorkbook returns center formatted XLWorkbook")]
        //public void GenerateWorkbookReturnsCenterFormattedXLWorkbook()
        //{
        //    var formatList = new List<ColumnModel>()
        //    {
        //        new ColumnModel
        //        {
        //            IsColumnCentered = true,
        //            ColumnName = "Service Date"
        //        }
        //    };

        //    var testWorkbook = subject
        //        .AddWorksheet(SampleDatesDataTable(), "test")
        //            .ConfigureColumn("Service Date", "Service Date", isCentered: true)
        //            .FinalizeWorksheet()
        //        .GenerateWorkbook();

        //    var streams = new Stream[]
        //    {
        //        CreateMemoryStream(testWorkbook),
        //        CreateMemoryStream(GenerateExpectedWorkbook(
        //            SampleDatesDataTable(),
        //            formatList))
        //    };

        //    excelEquality
        //        .StreamsAreEqual(streams)
        //        .AreEqual
        //        .ShouldBeTrue();
        //}

        //[Fact(DisplayName = "ReportGenerator: GenerateWorkbook returns black background formatted XLWorkbook")]
        //public void GenerateWorkbookReturnsBlackBackgroundFormattedXLWorkbook()
        //{
        //    var formatList = new List<ColumnModel>()
        //    {
        //        new ColumnModel
        //        {
        //            ColumnBackgroundColor = "#000000",
        //            ColumnName = "Service Date"
        //        }
        //    };

        //    var testWorkbook = subject
        //        .AddWorksheet(SampleDatesDataTable(), "test")
        //            .ConfigureColumn("Service Date", "Service Date", backgroundColor: "#000000")
        //            .FinalizeWorksheet()
        //        .GenerateWorkbook();

        //    var streams = new Stream[]
        //    {
        //        CreateMemoryStream(testWorkbook),
        //        CreateMemoryStream(GenerateExpectedWorkbook(
        //            SampleDatesDataTable(),
        //            formatList))
        //    };

        //    excelEquality
        //        .StreamsAreEqual(streams)
        //        .AreEqual
        //        .ShouldBeTrue();
        //}

        //[Fact(DisplayName = "ReportGenerator: GenerateWorkbook returns tan background formatted XLWorkbook")]
        //public void GenerateWorkbookReturnsTanBackgroundFormattedXLWorkbook()
        //{
        //    var formatList = new List<ColumnModel>()
        //    {
        //        new ColumnModel
        //        {
        //            ColumnBackgroundColor = "#FFFFCC",
        //            ColumnName = "Service Date"
        //        }
        //    };

        //    var testWorkbook = subject
        //        .AddWorksheet(SampleDatesDataTable(), "test")
        //            .ConfigureColumn("Service Date", "Service Date", backgroundColor: "#FFFFCC")
        //            .FinalizeWorksheet()
        //        .GenerateWorkbook();

        //    var streams = new Stream[]
        //    {
        //        CreateMemoryStream(testWorkbook),
        //        CreateMemoryStream(GenerateExpectedWorkbook(
        //            SampleDatesDataTable(),
        //            formatList))
        //    };

        //    excelEquality
        //        .StreamsAreEqual(streams)
        //        .AreEqual
        //        .ShouldBeTrue();
        //}

        [Fact(DisplayName = "ReportGenerator: ConfigureColumn throws human readable exception on mismatched name")]
        public void ConfigureColumnThrowsHumanReadableOnColumnMismatch()
        {
            Should.Throw<ArgumentException>(() =>
                subject
                    .AddWorksheet(SampleDataTable, "test")
                        .ConfigureColumn("totally wrong name", "new name")
                        .FinalizeWorksheet()
                    .GenerateWorkbook())
            .Message
            .ShouldBe("You must provide a matching DataColumn name as the first parameter.");
        }

        [Fact(DisplayName = "ReportGenerator: ConfigureColumn throws on invalid excel column name")]
        public void ConfigureColumnThrowsOnInvalidExcelColumnName()
        {
            Should.Throw<ArgumentException>(() =>
                subject
                    .AddWorksheet(SampleDataTable, "test")
                        .ConfigureColumn("Quantity", "")
                        .FinalizeWorksheet()
                    .GenerateWorkbook())
            .Message
            .ShouldBe("You must provide a valid Excel column name as the second parameter.");
        }

        [Fact(DisplayName = "ReportGenerator: Can add more than one worksheet")]
        public void CanAddMoreThanOneWorksheet()
        {
            var workbook = subject
                .AddWorksheet(SampleDataTable, "worksheet1")
                    .ConfigureColumn("Quantity", "Count")
                    .FinalizeWorksheet()
                .AddWorksheet(SampleDataTable, "worksheet2")
                    .ConfigureColumn("Item", "Food")
                    .ConfigureColumn("Person", "Dev", backgroundColor: "#FFFFCC")
                    .FinalizeWorksheet()
                .GenerateWorkbook();

            workbook.Worksheets.Count.ShouldBe(2);
        }
    }
}