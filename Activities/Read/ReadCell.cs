using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.Read
{
    [DisplayName("Read Cell")]
    [Description("Reads the value from a single cell in an Excel file.")]
    public class ReadCell : CodeActivity
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
        [DisplayName("3. Cell Address")]
        [Description("The address of the cell to read (e.g., \"A1\").")]
        public InArgument<string> CellAddress { get; set; }

        [Category("Options")]
        [DisplayName("4. Keep Formatting")]
        [Description("If true, reads the formatted string value. If false (default), reads the raw, typed value.")]
        [DefaultValue(false)]
        public InArgument<bool> KeepFormatting { get; set; } = false;

        [Category("Output")]
        [DisplayName("5. Value")]
        [Description("The value read from the cell.")]
        public OutArgument<object> Value { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var cellAddress = CellAddress.Get(context);
            var keepFormatting = KeepFormatting.Get(context);

            var result = Excel.ReadCell(filePath, sheetName, cellAddress, keepFormatting);
            Value.Set(context, result);
        }
    }
}