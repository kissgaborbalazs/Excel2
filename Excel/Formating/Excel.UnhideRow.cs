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
        /// Felfed egy elrejtett sort egy Excel fájlban.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap neve.</param>
        /// <param name="rowNumber">A felfedendő sor sorszáma.</param>
        public static void UnhideRow(string filePath, string sheetName, int rowNumber)
        {
            using (var manager = new ExcelManager(filePath))
            {
                manager.Worksheet(sheetName).UnhideRow(rowNumber);
            }
        }
    }
}
