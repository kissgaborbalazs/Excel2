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
        /// Rögzíti a munkalap paneljeit egy Excel fájlban.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap neve.</param>
        /// <param name="rowsToFreeze">A felülről rögzítendő sorok száma.</param>
        /// <param name="columnsToFreeze">A balról rögzítendő oszlopok száma.</param>
        public static void FreezePanes(string filePath, string sheetName, int rowsToFreeze, int columnsToFreeze)
        {
            using (var manager = new ExcelManager(filePath))
            {
                manager.Worksheet(sheetName).FreezePanes(rowsToFreeze, columnsToFreeze);
            }
        }
    }
}
