using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.Format
{
    [DisplayName("AutoFit Columns")]
    [Description("Adjusts the width of columns to fit their contents.")]
    public class AutoFitColumns : CodeActivity
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

        [Category("Options")]
        [DisplayName("3. Columns")]
        [Description("The column numbers to autofit (e.g., {1, 3, 5}). If empty, all columns will be adjusted.")]
        public InArgument<int[]> Columns { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var columns = Columns.Get(context) ?? new int[0];

            Excel.AutoFitColumns(filePath, sheetName, columns);
        }
    }
}