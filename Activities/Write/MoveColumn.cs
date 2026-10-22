using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.Data
{
    [DisplayName("Move Column")]
    [Description("Moves a column to a new position within the same worksheet.")]
    public class MoveColumn : CodeActivity
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
        [DisplayName("3. Source Column Number")]
        [Description("The number of the column to move.")]
        public InArgument<int> SourceColumnNumber { get; set; }

        [Category("Input")]
        [DisplayName("4. Destination Column Number")]
        [Description("The target position for the column.")]
        public InArgument<int> DestinationColumnNumber { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var sourceColumnNumber = SourceColumnNumber.Get(context);
            var destinationColumnNumber = DestinationColumnNumber.Get(context);

            Excel.MoveColumn(filePath, sheetName, sourceColumnNumber, destinationColumnNumber);
        }
    }
}