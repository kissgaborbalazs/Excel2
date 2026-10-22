using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.Read
{
    [DisplayName("Get Range Address")]
    [Description("Intelligently determines a contiguous data range and returns its address as a string.")]
    public class GetRangeAddress : CodeActivity
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
        [DisplayName("3. Start Cell")]
        [Description("The starting cell of the range. If empty, the entire used range is returned.")]
        public InArgument<string> StartCell { get; set; }

        [Category("Output")]
        [DisplayName("4. Range Address")]
        [Description("The address of the found range (e.g., \"A2:D15\").")]
        public OutArgument<string> RangeAddress { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var startCell = StartCell.Get(context);

            var result = Excel.GetRangeAddress(filePath, sheetName, startCell);
            RangeAddress.Set(context, result);
        }
    }
}