using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.Format
{
    [DisplayName("Set Number Format")]
    [Description("Sets the number format for a specified range (e.g., General, Number, Date, Text).")]
    public class SetNumberFormat : CodeActivity
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
        [Description("The range to format (e.g., \"A1:C10\").")]
        public InArgument<string> Range { get; set; }

        [Category("Options")]
        [DisplayName("4. Format Type")]
        [Description("Select a predefined format from the list. This is ignored if a Custom Format is provided.")]
        public NumberFormatType FormatType { get; set; }

        [Category("Options")]
        [DisplayName("5. Custom Format")]
        [Description("The Excel number format string. Overrides the Format Type. Examples: '#,##0.00', 'yyyy-mm-dd', '@'.")]
        public InArgument<string> CustomFormat { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var range = Range.Get(context);
            var customFormat = CustomFormat.Get(context);
            var formatType = FormatType;

            // Prioritási logika: Ha a szabadszöveges mező ki van töltve, azt használjuk.
            if (!string.IsNullOrWhiteSpace(customFormat))
            {
                Excel.SetNumberFormat(filePath, sheetName, range, customFormat);
            }
            else
            {
                // Ellenkező esetben a legördülő menüből választott enumot használjuk.
                Excel.SetNumberFormat(filePath, sheetName, range, formatType);
            }
        }
    }
}