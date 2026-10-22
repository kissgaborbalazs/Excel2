using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.Write
{
    [DisplayName("Set Formula")]
    [Description("Sets a formula for a specified cell.")]
    public class SetFormula : CodeActivity
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

        [Category("Input")]
        [DisplayName("4. Formula")]
        [Description("The formula to set, starting with '=' (e.g., \"=SUM(A1:A10)\").")]
        public InArgument<string> Formula { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var cellAddress = CellAddress.Get(context);
            var formula = Formula.Get(context);

            Excel.SetFormula(filePath, sheetName, cellAddress, formula);
        }
    }
}