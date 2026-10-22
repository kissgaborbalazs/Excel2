using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.Format
{
    [DisplayName("Unhide Row")]
    [Description("Unhides a specified row in a worksheet.")]
    public class UnhideRow : CodeActivity
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
        [Description("The number of the row to unhide.")]
        public InArgument<int> RowNumber { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var rowNumber = RowNumber.Get(context);

            Excel.UnhideColumn(filePath, sheetName, rowNumber);
        }
    }
}