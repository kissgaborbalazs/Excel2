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
        /// Visszaadja a szükséges oszlopnevet tartalmazó oszlop betűjelét.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap neve.</param>
        /// <param name="columnName">A keresendő oszlop neve</param>
        /// <returns>Az oszlop betűjele (pl. "D"), vagy null, ha a sor üres.</returns>
        public static string GetColumnLetter(string filePath, string sheetName, string columnName, int headerRow)
        {
            using (var manager = new ExcelManager(filePath))
            {
                return manager.Worksheet(sheetName).GetColumnLetter(columnName, headerRow);
            }
        }
    }
}
