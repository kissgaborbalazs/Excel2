using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.Format
{
    [DisplayName("Unhide Column")]
    [Description("Unhides a specified column in a worksheet.")]
    public class UnhideColumn : CodeActivity
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
        [DisplayName("3. Column Number")]
        [Description("The number of the column to unhide.")]
        public InArgument<int> ColumnNumber { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var columnNumber = ColumnNumber.Get(context);

            Excel.UnhideColumn(filePath, sheetName, columnNumber);
        }
    }
}