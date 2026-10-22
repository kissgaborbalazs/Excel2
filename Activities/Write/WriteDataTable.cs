using System.Activities;
using System.ComponentModel;
using System.Data;
using MPRT.Excel; // A saját logikai könyvtárunk importálása

namespace MPRT.Excel.Activities.Write
{
    [DisplayName("Write DataTable")]
    [Description("Writes a DataTable to a specified range in an Excel file without requiring Microsoft Office.")]
    public class WriteDataTable : CodeActivity
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

        [Category("Input")]
        [Description("The DataTable to be written to the Excel file.")]
        [DisplayName("3. Data Table")]
        public InArgument<DataTable> DataTable { get; set; }

        [Category("Input")]
        [Description("The starting cell for writing the data (e.g., \"A1\").")]
        [DisplayName("4. Start Cell")]
        public InArgument<string> StartCell { get; set; } = "A1";

        [Category("Options")]
        [Description("Specifies if the column names of the DataTable should be written as a header row.")]
        [DisplayName("5. Include Headers")]
        [DefaultValue(true)]
        public InArgument<bool> IncludeHeaders { get; set; } = true;


        /// <summary>
        /// Ez a metódus fut le, amikor a UiPath process futtatja az activity-t.
        /// </summary>
        protected override void Execute(CodeActivityContext context)
        {
            // 1. Bemeneti értékek kinyerése a kontextusból.
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var dataTable = DataTable.Get(context);
            var startCell = StartCell.Get(context);
            var includeHeaders = IncludeHeaders.Get(context);

            // 2. A logikai könyvtárunk meghívása.
            Excel.WriteDataTable(filePath, sheetName, dataTable, startCell, includeHeaders);
        }
    }
}