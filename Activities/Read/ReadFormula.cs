using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.Read
{
    [DisplayName("Read Formula")]
    [Description("Reads the formula from a specified cell as a string.")]
    public class ReadFormula : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [DisplayName("1. File Path")]
        [Description("The full path to the Excel file.")]
        public InArgument<string> FilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("2. Sheet Name")]
        public InArgument<string> SheetName { get; set; }

        [Category("Input")]
        [DisplayName("3. Cell Address")]
        public InArgument<string> CellAddress { get; set; }

        [Category("Output")]
        [DisplayName("4. Formula")]
        [Description("The formula from the cell as a string.")]
        public OutArgument<string> Formula { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var cellAddress = CellAddress.Get(context);

            var result = Excel.ReadFormula(filePath, sheetName, cellAddress);
            Formula.Set(context, result);
        }
    }
}