using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.File
{
    [DisplayName("Clear Sheet")]
    [Description("Clears the content and optionally the formatting of an entire worksheet.")]
    public class ClearSheet : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [DisplayName("1. File Path")]
        [Description("The full path to the Excel file.")]
        public InArgument<string> FilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("2. Sheet Name")]
        [Description("The name of the worksheet to clear.")]
        public InArgument<string> SheetName { get; set; }

        [Category("Options")]
        [DisplayName("3. Clear Styles")]
        [Description("If true, clears formatting as well as content. If false (default), only clears content.")]
        [DefaultValue(false)]
        public bool ClearStyles { get; set; } = false;

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var clearStyles = ClearStyles;

            Excel.ClearSheet(filePath, sheetName, clearStyles);
        }
    }
}