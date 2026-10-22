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
        /// Kiolvassa egy cella képletét egy Excel fájlból.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap neve.</param>
        /// <param name="cellAddress">A cella címe.</param>
        /// <returns>A cella képlete stringként.</returns>
        public static string ReadFormula(string filePath, string sheetName, string cellAddress)
        {
            using (var manager = new ExcelManager(filePath))
            {
                return manager.Worksheet(sheetName).ReadFormula(cellAddress);
            }
        }
    }
}
