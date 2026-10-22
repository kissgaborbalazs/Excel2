using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.Read
{
    [DisplayName("Get Letter Of Specific Column")]
    [Description("Gets the column letter (e.g. \"A\", \"AB\") of the given column name.")]
    public class GetColumnLetter : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [DisplayName("1. File Path")]
        [Description("The full path to the Excel file.")]
        public InArgument<string> FilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("2. Sheet Name")]
        [Description("The name of the worksheet.")]
        public InArgument<string> SheetName { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("3. Column Name")]
        [Description("The name of the column to inspect (e.g. \"A\", \"AB\").")]
        public InArgument<string> ColumnName { get; set; }
        
        [Category("Input")]
        [DisplayName("4. Header Row")]
        [DefaultValue(1)]
        [Description("The row number (not index!), that contains the column to find (1 by default)")]
        public InArgument<int> HeaderRow { get; set; } = 1;

        [Category("Output")]
        [DisplayName("5. Column Letter")]
        [Description("The letter searched column name.")]
        public OutArgument<string> ColumnLetter { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var columnName = ColumnName.Get(context);
            var headerRow = HeaderRow.Get(context);

            var result = Excel.GetColumnLetter(filePath, sheetName, columnName, headerRow);
            ColumnLetter.Set(context, result);
        }
    }
}