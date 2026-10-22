using System.Activities;
using System.ComponentModel;
using System.Data;
using MPRT.Excel; // A saját logikai könyvtárunk importálása

namespace MPRT.Excel.Activities.Read
{
    [DisplayName("Read Range")]
    [Description("Reads a range from an Excel file into a DataTable without requiring Microsoft Office.")]
    public class ReadRange : CodeActivity
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
        [Description("The range to read (e.g., \"A1:C10\"). If empty, reads the entire used range. If a single cell, expands to the contiguous data region.")]
        [DisplayName("3. Range")]
        public InArgument<string> Range { get; set; }

        [Category("Options")]
        [Description("Specifies if the first row of the range is a header.")]
        [DisplayName("Has Header")]
        [DefaultValue(true)]
        public bool HasHeader { get; set; } = true;

        [Category("Options")]
        [Description("If true, reads the formatted string value. If false (default), reads the raw, typed value (e.g., double, DateTime).")]
        [DisplayName("Keep Formatting")]
        [DefaultValue(false)]
        public bool KeepFormatting { get; set; } = false;

        // --- Kimeneti tulajdonság (Output Property) ---
        // Az OutArgument<T> típus jelzi, hogy ez egy kimeneti mező, ahova az eredményt mentjük.

        [Category("Output")]
        [Description("The DataTable containing the read data.")]
        [DisplayName("Data Table")]
        public OutArgument<DataTable> DataTable { get; set; }

        /// <summary>
        /// Ez a metódus fut le, amikor a UiPath process futtatja az activity-t.
        /// </summary>
        protected override void Execute(CodeActivityContext context)
        {
            // 1. Bemeneti értékek kinyerése a kontextusból.
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var range = Range.Get(context);
            var hasHeader = HasHeader;
            var keepFormatting = KeepFormatting;

            // 2. A logikai könyvtárunk meghívása.
            var resultDataTable = Excel.ReadRange(filePath, sheetName, range, hasHeader, keepFormatting);

            // 3. A kimeneti érték beállítása.
            // A Set metódussal töltjük fel az OutArgument-et az eredménnyel.
            DataTable.Set(context, resultDataTable);
        }
    }
}