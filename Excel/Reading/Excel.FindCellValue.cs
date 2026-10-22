namespace MPRT.Excel
{
    //// <summary>
    /// Egyszerű, statikus segédmetódusokat biztosít a leggyakoribb Excel műveletekhez.
    /// Minden metódus egy teljes, önálló műveletet hajt végre (megnyitás, módosítás, mentés, bezárás).
    /// Akkor ideális, ha csak egyetlen műveletet kell elvégezni egy fájlon.
    /// </summary>
    public static partial class Excel
    {
        /// <summary>
        /// Megkeresi egy érték első előfordulását egy munkalapon, és visszaadja a cella címét.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap neve.</param>
        /// <param name="valueToFind">A keresendő érték.</param>
        /// <returns>A megtalált cella címe (pl. "A5"), vagy üres string, ha nem található.</returns>
        public static string FindCellValue(string filePath, string sheetName, object valueToFind)
        {
            // Cél: Egyszerűség és kényelem biztosítása a kereséshez.
            // A 'using' blokk garantálja a biztonságos fájlkezelést (a fájl mindig
            // bezárásra kerül, még hiba esetén is).
            using (var manager = new ExcelManager(filePath))
            {
                // A tényleges munkát a WorksheetExtensions-ben definiált, láncolható FindCell metódusra delegálja.
                return manager.Worksheet(sheetName).FindCell(valueToFind);
            }
        }
    }
}