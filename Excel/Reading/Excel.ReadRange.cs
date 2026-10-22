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
        /// Beolvas egy megadott tartományt egy DataTable-be egyetlen, atomi művelettel.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap neve.</param>
        /// <param name="range">Az olvasandó tartomány (pl. "A1:C10"). Ha null, a teljes használt területet olvassa.</param>
        /// <param name="hasHeader">Igaz, ha a tartomány első sora a fejléc.</param>
        /// <param name="keepFormatting">Igaz, ha a cella formázott, látható értékét (string) kell beolvasni. Hamis (alapértelmezett), ha a nyers, típusos adatot (double, DateTime, stb.).</param>
        /// <returns>Egy DataTable a beolvasott adatokkal.</returns>
        public static DataTable ReadRange(string filePath, string sheetName, string range = null, bool hasHeader = true, bool keepFormatting = false)
        {
            // Cél: Egyszerűség és kényelem biztosítása.
            // Ez a metódus elrejti a belső működés komplexitását (fájl megnyitása, munkalap kezelése, bezárás).
            // Létrehoz egy 'ExcelManager' példányt egy 'using' blokkban, ami garantálja a biztonságos
            // fájlkezelést (a fájl mindig mentésre és bezárásra kerül, még hiba esetén is).
            using (var manager = new ExcelManager(filePath))
            {
                // A tényleges munkát a 'WorksheetExtensions'-ben definiált, láncolható ReadRange metódusra delegálja.
                // Ez a kód-újrahasznosítás legjobb gyakorlatát követi, elkerülve a logika duplikálását.
                return manager.Worksheet(sheetName).ReadRange(range, hasHeader, keepFormatting);
            }
        }
    }
}