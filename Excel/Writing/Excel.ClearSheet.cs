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
        /// Törli egy teljes munkalap tartalmát egy Excel fájlban.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A törlendő munkalap neve.</param>
        /// <param name="clearStyles">Igaz, ha a formázást is törölni kell.</param>
        public static void ClearSheet(string filePath, string sheetName, bool clearStyles = false)
        {
            using (var manager = new ExcelManager(filePath))
            {
                manager.Worksheet(sheetName).Clear(clearStyles);
            }
        }
    }
}
