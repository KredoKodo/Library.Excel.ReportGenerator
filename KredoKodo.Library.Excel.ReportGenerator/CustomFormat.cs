namespace KredoKodo.Library.Excel.ReportGenerator
{
    /// <summary>
    /// Represents a custom column formatting.
    /// </summary>
    public static class CustomFormat
    {
        /// <summary>
        /// Numbers formatted with a comma as the thousands separator.
        /// format: #,##0
        /// </summary>
        /// <returns>#,##0</returns>
        public static string Quantity
        {
            get
            {
                return "#,##0";
            }
        }
        /// <summary>
        /// Short Currency (No decimal).
        /// format: $#,##0
        /// </summary>
        /// <returns>$#,##0</returns>
        public static string Currency_0_Decimals
        {
            get
            {
                return "$#,##0";
            }
        }
        /// <summary>
        /// Standard Currency (2 digits after the decimal).
        /// format: $#,##0.00
        /// </summary>
        /// <returns>$#,##0.00</returns>
        public static string Currency_2_Decimals
        {
            get
            {
                return "$#,##0.00";
            }
        }
        /// <summary>
        /// Long Currency (4 digits after the decimal).
        /// format: $#,##0.0000
        /// </summary>
        /// <returns>$#,##0.0000</returns>
        public static string Currency_4_Decimals
        {
            get
            {
                return "$#,##0.0000";
            }
        }
        /// <summary>
        /// Short date in this format: mm/dd/yyyy
        /// </summary>
        /// <returns>mm/dd/yyyy</returns>
        public static string ShortDate
        {
            get
            {
                return "mm/dd/yyyy";
            }
        }
    }
}