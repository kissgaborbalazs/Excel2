using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.File
{
    [DisplayName("Copy Sheet To Workbook")]
    [Description("Copies a worksheet from a source Excel file to a destination Excel file.")]
    public class CopySheetToWorkbook : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [DisplayName("1. Source File Path")]
        [Description("The full path to the source Excel file.")]
        public InArgument<string> SourceFilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("2. Source Sheet Name")]
        [Description("The name of the worksheet to copy from the source file.")]
        public InArgument<string> SourceSheetName { get; set; }

        [Category("Input")]
        [DisplayName("3. Destination File Path")]
        [Description("The full path to the destination Excel file. It will be created if it does not exist.")]
        public InArgument<string> DestinationFilePath { get; set; }

        [Category("Options")]
        [DisplayName("4. Destination Sheet Name")]
        [Description("The name for the copied worksheet in the destination file.")]
        public InArgument<string> DestinationSheetName { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var sourceFilePath = SourceFilePath.Get(context);
            var sourceSheetName = SourceSheetName.Get(context);
            var destinationFilePath = DestinationFilePath.Get(context);
            var destinationSheetName = DestinationSheetName.Get(context);

            Excel.CopySheetToWorkbook(sourceFilePath, sourceSheetName, destinationFilePath, destinationSheetName);
        }
    }
}