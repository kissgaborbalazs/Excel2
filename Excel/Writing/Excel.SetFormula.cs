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
        /// Képletet állít be egy cellához egy Excel fájlban.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap neve.</param>
        /// <param name="cellAddress">A cella címe.</param>
        /// <param name="formula">A beírandó képlet (pl. "=SUM(A1:A10)").</param>
        public static void SetFormula(string filePath, string sheetName, string cellAddress, string formula)
        {
            using (var manager = new ExcelManager(filePath))
            {
                manager.Worksheet(sheetName).SetFormula(cellAddress, formula);
            }
        }
    }
}
