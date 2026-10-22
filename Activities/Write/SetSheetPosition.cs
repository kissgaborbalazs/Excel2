using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.Write
{
    [DisplayName("Set Sheet Position")]
    [Description("Moves a worksheet to a specific position (1-based).")]
    public class SetSheetPosition : CodeActivity
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

        [Category("Input")]
        [DisplayName("3. Position")]
        [Description("The target position for the worksheet (1-based index).")]
        public InArgument<int> Position { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var position = Position.Get(context);

            Excel.SetSheetPosition(filePath, sheetName, position);
        }
    }
}