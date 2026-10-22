using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.Read
{
    [DisplayName("Get Last Row In Column")]
    [Description("Gets the number of the last row containing data in a specific column.")]
    public class GetLastRowInColumn : CodeActivity
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
        [DisplayName("3. Column Letter")]
        [Description("The letter of the column to inspect (e.g., \"A\").")]
        public InArgument<string> ColumnLetter { get; set; }

        [Category("Output")]
        [DisplayName("4. Last Row Number")]
        [Description("The number of the last row with data in the specified column.")]
        public OutArgument<int> LastRowNumber { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var columnLetter = ColumnLetter.Get(context);

            var result = Excel.GetLastRow(filePath, sheetName, columnLetter);
            LastRowNumber.Set(context, result);
        }
    }
}