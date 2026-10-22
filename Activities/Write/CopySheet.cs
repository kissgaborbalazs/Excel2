using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.File
{
    [DisplayName("Copy Sheet")]
    [Description("Copies a worksheet within the same Excel file.")]
    public class CopySheet : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [DisplayName("1. File Path")]
        [Description("The full path to the Excel file.")]
        public InArgument<string> FilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("2. Source Sheet Name")]
        [Description("The name of the worksheet to copy.")]
        public InArgument<string> SourceSheetName { get; set; }

        [Category("Input")]
        [DisplayName("3. New Sheet Name")]
        [Description("The name for the newly created copy.")]
        public InArgument<string> NewSheetName { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sourceSheetName = SourceSheetName.Get(context);
            var newSheetName = NewSheetName.Get(context);

            Excel.CopySheet(filePath, sourceSheetName, newSheetName);
        }
    }
}