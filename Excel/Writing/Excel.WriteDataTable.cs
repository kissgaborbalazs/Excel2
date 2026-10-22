using System.Data;

namespace MPRT.Excel
{
    /// <summary>
    /// Egyszerű, statikus segédmetódusokat biztosít a leggyakoribb Excel műveletekhez.
    /// Minden metódus egy teljes, önálló műveletet hajt végre (megnyitás, módosítás, mentés, bezárás).
    /// Akkor ideális, ha csak egyetlen műveletet kell elvégezni egy fájlon.
    /// </summary>
    public static partial class Excel
    {
        /// <summary>
        /// Beír egy teljes DataTable-t a megadott munkalapra egyetlen, atomi művelettel.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap neve.</param>
        /// <param name="dataTable">A beírandó adatok.</param>
        /// <param name="startCell">A cella, ahonnan az írás kezdődik (pl. "A1").</param>
        /// <param name="includeHeaders">Igaz, ha a DataTable oszlopneveit is be kell írni fejlécként.</param>
        public static void WriteDataTable(string filePath, string sheetName, DataTable dataTable, string startCell = "A1", bool includeHeaders = true)
        {
            // Cél: Egyszerűség és kényelem biztosítása.
            // Ez a metódus elrejti a belső működés komplexitását (fájl megnyitása, munkalap kezelése, bezárás).
            // Létrehoz egy 'ExcelManager' példányt egy 'using' blokkban, ami garantálja a biztonságos
            // fájlkezelést (a fájl mindig mentésre és bezárásra kerül, még hiba esetén is).
            using (var manager = new ExcelManager(filePath))
            {
                // A tényleges munkát a 'WorksheetExtensions'-ben definiált, láncolható WriteDataTable metódusra delegálja.
                // Ez a kód-újrahasznosítás legjobb gyakorlatát követi, elkerülve a logika duplikálását.
                manager.Worksheet(sheetName).WriteDataTable(dataTable, startCell, includeHeaders);
            }
        }
    }
}