using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.Data
{
    [DisplayName("Move Row")]
    [Description("Moves a row to a new position within the same worksheet.")]
    public class MoveRow : CodeActivity
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
        [DisplayName("3. Source Row Number")]
        [Description("The number of the row to move.")]
        public InArgument<int> SourceRowNumber { get; set; }

        [Category("Input")]
        [DisplayName("4. Destination Row Number")]
        [Description("The target position for the row.")]
        public InArgument<int> DestinationRowNumber { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var sourceRowNumber = SourceRowNumber.Get(context);
            var destinationRowNumber = DestinationRowNumber.Get(context);

            Excel.MoveRow(filePath, sheetName, sourceRowNumber, destinationRowNumber);
        }
    }
}
