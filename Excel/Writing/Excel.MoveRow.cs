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
        /// Áthelyez egy sort egy új pozícióba egy Excel fájlban.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap neve.</param>
        /// <param name="sourceRowNumber">A mozgatandó sor sorszáma.</param>
        /// <param name="destinationRowNumber">A célpozíció sorszáma.</param>
        public static void MoveRow(string filePath, string sheetName, int sourceRowNumber, int destinationRowNumber)
        {
            using (var manager = new ExcelManager(filePath))
            {
                manager.Worksheet(sheetName).MoveRow(sourceRowNumber, destinationRowNumber);
            }
        }
    }
}
