using System.Activities;
using System.ComponentModel;
using MPRT.Excel; // A saját logikai könyvtárunk importálása

namespace MPRT.Excel.Activities.Read
{
    [DisplayName("Find Cell Value")]
    [Description("Finds the first occurrence of a value in a worksheet and returns its cell address.")]
    public class FindCellValue : CodeActivity
    {
        // --- Bemeneti tulajdonságok (Input Properties) ---

        [Category("Input")]
        [RequiredArgument]
        [Description("The full path to the Excel file.")]
        [DisplayName("1. File Path")]
        public InArgument<string> FilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("The name of the worksheet to search in.")]
        [DisplayName("2. Sheet Name")]
        public InArgument<string> SheetName { get; set; }

        [Category("Input")]
        [Description("The value to find in the worksheet.")]
        [DisplayName("3. Value to Find")]
        public InArgument<object> ValueToFind { get; set; }

        // --- Kimeneti tulajdonság (Output Property) ---

        [Category("Output")]
        [Description("The address of the found cell (e.g., \"A5\"). Returns empty if not found.")]
        [DisplayName("4. Cell Address")]
        public OutArgument<string> CellAddress { get; set; }


        /// <summary>
        /// Ez a metódus fut le, amikor a UiPath process futtatja az activity-t.
        /// </summary>
        protected override void Execute(CodeActivityContext context)
        {
            // 1. Bemeneti értékek kinyerése a kontextusból.
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var valueToFind = ValueToFind.Get(context);

            // 2. A logikai könyvtárunk meghívása.
            var foundAddress = Excel.FindCellValue(filePath, sheetName, valueToFind);

            // 3. A kimeneti érték beállítása.
            CellAddress.Set(context, foundAddress);
        }
    }
}