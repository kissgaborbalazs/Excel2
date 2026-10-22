using System.Activities;
using System.ComponentModel;
using System.Data;
using MPRT.Excel; // A saját logikai könyvtárunk importálása

namespace MPRT.Excel.Activities.Write
{
    [DisplayName("Append DataTable")]
    [Description("Appends a DataTable to the existing data in a worksheet without requiring Microsoft Office.")]
    public class AppendDataTable : CodeActivity
    {
        // --- Bemeneti tulajdonságok (Input Properties) ---

        [Category("Input")]
        [RequiredArgument]
        [Description("The full path to the Excel file.")]
        [DisplayName("1. File Path")]
        public InArgument<string> FilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("The name of the worksheet to append to.")]
        [DisplayName("2. Sheet Name")]
        public InArgument<string> SheetName { get; set; }

        [Category("Input")]
        [Description("The DataTable to be appended.")]
        [DisplayName("3. Data Table")]
        public InArgument<DataTable> DataTable { get; set; }


        /// <summary>
        /// Ez a metódus fut le, amikor a UiPath process futtatja az activity-t.
        /// </summary>
        protected override void Execute(CodeActivityContext context)
        {
            // 1. Bemeneti értékek kinyerése a kontextusból.
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var dataTable = DataTable.Get(context);

            // 2. A logikai könyvtárunk meghívása.
            Excel.AppendDataTable(filePath, sheetName, dataTable);
        }
    }
}