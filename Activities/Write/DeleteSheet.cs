using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.Write
{
    [DisplayName("Delete Sheet")]
    [Description("Deletes a worksheet from an Excel file.")]
    public class DeleteSheet : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [DisplayName("1. File Path")]
        [Description("The full path to the Excel file.")]
        public InArgument<string> FilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("2. Sheet Name")]
        [Description("The name of the worksheet to delete.")]
        public InArgument<string> SheetName { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);

            Excel.DeleteSheet(filePath, sheetName);
        }
    }
}
