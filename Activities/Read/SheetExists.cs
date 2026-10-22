using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.Read
{
    [DisplayName("Sheet Exists")]
    [Description("Checks if a worksheet with a specified name exists in the Excel file.")]
    public class SheetExists : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [DisplayName("1. File Path")]
        [Description("The full path to the Excel file.")]
        public InArgument<string> FilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("2. Sheet Name")]
        [Description("The name of the worksheet to check.")]
        public InArgument<string> SheetName { get; set; }

        [Category("Output")]
        [DisplayName("3. Exists")]
        [Description("Returns true if the worksheet exists, otherwise false.")]
        public OutArgument<bool> Exists { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var result = Excel.SheetExists(filePath, sheetName);
            Exists.Set(context, result);
        }
    }
}
