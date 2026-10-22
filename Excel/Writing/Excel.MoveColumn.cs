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
        /// Áthelyez egy oszlopot egy új pozícióba egy Excel fájlban.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap neve.</param>
        /// <param name="sourceColumnNumber">A mozgatandó oszlop sorszáma.</param>
        /// <param name="destinationColumnNumber">A célpozíció sorszáma.</param>
        public static void MoveColumn(string filePath, string sheetName, int sourceColumnNumber, int destinationColumnNumber)
        {
            using (var manager = new ExcelManager(filePath))
            {
                manager.Worksheet(sheetName).MoveColumn(sourceColumnNumber, destinationColumnNumber);
            }
        }
    }
}