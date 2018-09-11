namespace KredoKodo.Library.Excel.ReportGenerator
{
    /// <summary>
    /// Represents one column with customization applied
    /// </summary>
    public class ColumnModel
    {
        /// <summary>
        /// Initializes a new instance of the
        /// KredoKodo.Library.Excel.ReportGenerator.ColumnModel class.
        /// </summary>
        public ColumnModel()
        {
            IsColumnCentered = false;
        }

        /// <summary>
        /// Gets or sets the name of the column.
        /// </summary>
        public string ColumnName { get; set; }

        /// <summary>
        /// Gets or sets the formatting of the column.
        /// </summary>
        public string ColumnFormat { get; set; }

        /// <summary>
        /// Gets or sets whether column is centered.
        /// </summary>
        public bool IsColumnCentered { get; set; }

        /// <summary>
        /// Gets or sets the background color of the column.
        /// </summary>
        public string ColumnBackgroundColor { get; set; }
    }
}