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
        /// A megadott oszlopok szélességét a tartalmukhoz igazítja egy Excel fájlban.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap neve.</param>
        /// <param name="columns">Az igazítandó oszlopok sorszámai. Ha üres, az összes oszlopot igazítja.</param>
        public static void AutoFitColumns(string filePath, string sheetName, params int[] columns)
        {
            using (var manager = new ExcelManager(filePath))
            {
                manager.Worksheet(sheetName).AutoFitColumns(columns);
            }
        }
    }
}