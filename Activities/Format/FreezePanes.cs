using System.Activities;
using System.ComponentModel;
using MPRT.Excel; // A saját logikai könyvtárunk importálása

namespace MPRT.Excel.Activities.Format
{
    [DisplayName("Freeze Panes")]
    [Description("Freezes the top rows and/or left columns of a worksheet.")]
    public class FreezePanes : CodeActivity
    {
        // --- Bemeneti tulajdonságok (Input Properties) ---

        [Category("Input")]
        [RequiredArgument]
        [Description("The full path to the Excel file.")]
        [DisplayName("1. File Path")]
        public InArgument<string> FilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("The name of the worksheet.")]
        [DisplayName("2. Sheet Name")]
        public InArgument<string> SheetName { get; set; }

        [Category("Options")]
        [Description("The number of rows to freeze from the top. Set to 0 to not freeze any rows.")]
        [DisplayName("3. Rows to Freeze")]
        [DefaultValue(0)]
        public InArgument<int> RowsToFreeze { get; set; } = 1;

        [Category("Options")]
        [Description("The number of columns to freeze from the left. Set to 0 to not freeze any columns.")]
        [DisplayName("4. Columns to Freeze")]
        [DefaultValue(0)]
        public InArgument<int> ColumnsToFreeze { get; set; } = 0;


        /// <summary>
        /// Ez a metódus fut le, amikor a UiPath process futtatja az activity-t.
        /// </summary>
        protected override void Execute(CodeActivityContext context)
        {
            // 1. Bemeneti értékek kinyerése a kontextusból.
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var rowsToFreeze = RowsToFreeze.Get(context);
            var columnsToFreeze = ColumnsToFreeze.Get(context);

            // 2. A logikai könyvtárunk meghívása.
            Excel.FreezePanes(filePath, sheetName, rowsToFreeze, columnsToFreeze);
        }
    }
}