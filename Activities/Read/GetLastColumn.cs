using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.Read
{
    [DisplayName("Get Last Column Letter In Row")]
    [Description("Gets the column letter of the last cell containing data in a specific row.")]
    public class GetLastColumnLetterInRow : CodeActivity
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
        [DisplayName("3. Row Number")]
        [Description("The number of the row to inspect.")]
        public InArgument<int> RowNumber { get; set; }

        [Category("Output")]
        [DisplayName("4. Column Letter")]
        [Description("The column letter of the last used cell in the row (e.g., \"D\").")]
        public OutArgument<string> ColumnLetter { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var rowNumber = RowNumber.Get(context);

            var result = Excel.GetLastColumn(filePath, sheetName, rowNumber);
            ColumnLetter.Set(context, result);
        }
    }
}
