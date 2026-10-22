using System;
using System.Activities;
using System.ComponentModel;
using ClosedXML.Excel; // Szükséges az XLSortOrder enum-hoz
using MPRT.Excel; // A saját logikai könyvtárunk importálása

namespace MPRT.Excel.Activities.Format
{
    [DisplayName("Sort Range")]
    [Description("Sorts a specified range in an Excel file based on a column.")]
    public class SortRange : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("The full path to the Excel file.")]
        [DisplayName("1. File Path")]
        public InArgument<string> FilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("The name of the worksheet.")]
        [DisplayName("2. Sheet Name")]
        public InArgument<string> SheetName { get; set; }

        [Category("Input")]
        [Description("The range to be sorted (e.g., \"A1:C10\").")]
        [DisplayName("3. Range")]
        public InArgument<string> Range { get; set; }

        [Category("Input")]
        [Description("The column to sort by. Can be a column number (e.g., 1) or a column name (e.g., \"ID\").")]
        [DisplayName("4a. Sort By Column Number")]
        public InArgument<int?> ColumnIndex { get; set; }

        [Category("Input")]
        [Description("The column to sort by. Can be a column number (e.g., 1) or a column name (e.g., \"ID\").")]
        [DisplayName("4b. Sort By Column Letter")]
        public InArgument<string> ColumnLetter { get; set; }

        [Category("Input")]
        [Description("The column to sort by. Can be a column number (e.g., 1) or a column name (e.g., \"ID\").")]
        [DisplayName("4c. Sort By Column Name")]
        public InArgument<string> ColumnName { get; set; }

        [Category("Options")]
        [DisplayName("5. Has Header")]
        [Description("Specifies if the range has a header row that should be excluded from sorting.")]
        [DefaultValue(true)]
        public bool HasHeader { get; set; } = true;

        [Category("Options")]
        [Description("The sort order (Ascending or Descending).")]
        [DisplayName("6. Sort Order")]
        public XLSortOrder SortOrder { get; set; } = XLSortOrder.Ascending;


        /// <summary>
        /// Ez a metódus fut le, amikor a UiPath process futtatja az activity-t.
        /// </summary>
        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var range = Range.Get(context);
            var hasHeader = HasHeader;
            var sortOrder = SortOrder;

            var columnIndex = ColumnIndex.Get(context);
            var columnLetter = ColumnLetter.Get(context);
            var columnName = ColumnName.Get(context);

            if (columnIndex.HasValue)
            {
                Excel.SortRange(filePath, sheetName, range, columnIndex.Value, sortOrder, hasHeader);
            }
            else if (!string.IsNullOrWhiteSpace(columnLetter))
            {
                Excel.SortRange(filePath, sheetName, range, columnLetter, sortOrder, hasHeader);
            }
            else if (!string.IsNullOrWhiteSpace(columnName))
            {
                Excel.SortRangeByName(filePath, sheetName, range, columnName, sortOrder);
            }
            else
            {
                throw new ArgumentOutOfRangeException("One of the following sort criteria must be provided: Column Index, Column Letter, or Column Name.");
            }
        }
    }
}