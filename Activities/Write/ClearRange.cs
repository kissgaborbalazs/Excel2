using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.Write
{
    [DisplayName("Clear Range")]
    [Description("Clears the content and optionally the formatting of a specified range.")]
    public class ClearRange : CodeActivity
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
        [DisplayName("3. Range")]
        [Description("The range to clear (e.g., \"A1:D10\").")]
        public InArgument<string> Range { get; set; }

        [Category("Options")]
        [DisplayName("4. Clear Styles")]
        [Description("If true, clears formatting as well as content. If false (default), only clears content.")]
        [DefaultValue(false)]
        public bool ClearStyles { get; set; } = false;

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var range = Range.Get(context);
            var clearStyles = ClearStyles;

            Excel.ClearRange(filePath, sheetName, range, clearStyles);
        }
    }
}