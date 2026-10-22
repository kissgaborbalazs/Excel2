using System.Activities;
using System.ComponentModel;

using MPRT.Excel; // A saját logikai könyvtárunk importálása

namespace MPRT.Excel.Activities.Write
{
    // A DisplayName attribútum adja meg az Activity nevét a UiPath Studio-ban.
    [DisplayName("Write Cell")]
    // A Description adja a leírást, ami megjelenik, ha az Activity fölé visszük az egeret.
    [Description("Writes a value to a specified cell in an Excel file without requiring Microsoft Office.")]
    public class WriteCell : CodeActivity
    {
        // --- Bemeneti tulajdonságok (Input Properties) ---
        // Az InArgument<T> típus jelzi, hogy ez egy bemeneti mező az Activity Properties paneljén.
        // A Category attribútummal csoportosíthatjuk a tulajdonságokat.
        // A RequiredArgument attribútum kötelezővé teszi a mező kitöltését.

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
        [Description("The address of the cell (e.g., \"A1\").")]
        [DisplayName("3. Cell Address")]
        public InArgument<string> CellAddress { get; set; }

        [Category("Input")]
        [DisplayName("4. Value")]
        [Description("The value to be written into the cell.")]
        public InArgument<object> Value { get; set; }

        /// <summary>
        /// Ez a metódus fut le, amikor a UiPath process futtatja az activity-t.
        /// </summary>
        /// <param name="context">A futási környezet, amiből ki tudjuk nyerni a bemeneti értékeket.</param>
        protected override void Execute(CodeActivityContext context)
        {
            // 1. Bemeneti értékek kinyerése a kontextusból.
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var cellAddress = CellAddress.Get(context);
            var value = Value.Get(context);

            // 2. A logikai könyvtárunk meghívása.
            // Cél: Az Activity "code-behind"-ja maradjon vékony és tiszta.
            // A tényleges, komplex logikát a már letesztelt MPRT.Excel.Excel osztály végzi.
            // Ez a tiszta szétválasztás (separation of concerns) legjobb gyakorlatát követi.
            Excel.WriteCell(filePath, sheetName, cellAddress, value);
        }
    }
}