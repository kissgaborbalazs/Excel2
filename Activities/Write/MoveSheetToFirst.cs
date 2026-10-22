using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.File
{
    [DisplayName("Move Sheet To First")]
    [Description("Moves a worksheet to the first position.")]
    public class MoveSheetToFirst : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [DisplayName("1. File Path")]
        [Description("The full path to the Excel file.")]
        public InArgument<string> FilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("2. Sheet Name")]
        [Description("The name of the worksheet to move.")]
        public InArgument<string> SheetName { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);

            Excel.MoveSheetToFirst(filePath, sheetName);
        }
    }
}