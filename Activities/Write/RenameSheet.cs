using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.Write
{
    [DisplayName("Rename Sheet")]
    [Description("Renames a worksheet in an Excel file.")]
    public class RenameSheet : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [DisplayName("1. File Path")]
        [Description("The full path to the Excel file.")]
        public InArgument<string> FilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("2. Sheet Name")]
        [Description("The current name of the worksheet.")]
        public InArgument<string> SheetName { get; set; }

        [Category("Input")]
        [DisplayName("3. New Sheet Name")]
        [Description("The new name for the worksheet.")]
        public InArgument<string> NewSheetName { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var newSheetName = NewSheetName.Get(context);

            Excel.RenameSheet(filePath, sheetName, newSheetName);
        }
    }
}
