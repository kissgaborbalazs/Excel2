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
        /// Ellenőrzi, hogy egy adott nevű munkalap létezik-e egy Excel fájlban.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A keresendő munkalap neve.</param>
        /// <returns>Igaz (true), ha a munkalap létezik; egyébként hamis (false).</returns>
        public static bool SheetExists(string filePath, string sheetName)
        {
            // Cél: Egyszerűség és kényelem biztosítása.
            // A 'using' blokk garantálja a biztonságos fájlkezelést.
            using (var manager = new ExcelManager(filePath))
            {
                // A tényleges munkát az ExcelManager-ben definiált metódusra delegálja.
                return manager.SheetExists(sheetName);
            }
        }
    }
}