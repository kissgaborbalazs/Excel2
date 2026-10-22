using System.Activities;
using System.ComponentModel;
using System.Collections.Generic;
using MPRT.Excel;

namespace MPRT.Excel.Activities.Read
{
    [DisplayName("Get Sheet Names")]
    [Description("Retrieves a list of all worksheet names from an Excel file.")]
    public class GetSheetNames : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [DisplayName("1. File Path")]
        [Description("The full path to the Excel file.")]
        public InArgument<string> FilePath { get; set; }

        [Category("Output")]
        [DisplayName("2. Sheet Names")]
        [Description("A list of strings containing the names of the worksheets.")]
        public OutArgument<List<string>> SheetNames { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var result = Excel.GetSheetNames(filePath);
            SheetNames.Set(context, result);
        }
    }
}