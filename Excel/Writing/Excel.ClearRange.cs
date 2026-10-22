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
        /// Törli egy megadott tartomány tartalmát egy Excel fájlban.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap neve.</param>
        /// <param name="rangeAddress">A törlendő tartomány címe (pl. "A1:D10").</param>
        /// <param name="clearStyles">Igaz, ha a formázást is törölni kell.</param>
        public static void ClearRange(string filePath, string sheetName, string rangeAddress, bool clearStyles = false)
        {
            using (var manager = new ExcelManager(filePath))
            {
                manager.Worksheet(sheetName).Range(rangeAddress).Clear(clearStyles);
            }
        }
    }
}
