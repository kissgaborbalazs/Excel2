using ClosedXML.Excel;

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
        /// Értéket ír egy megadott cellába egyetlen, atomi művelettel.
        /// A háttérben az ExcelManager-t használja a robusztus fájlkezeléshez.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap neve.</param>
        /// <param name="cellAddress">A cella címe (pl. "A1").</param>
        /// <param name="value">A cellába írandó érték.</param>
        public static void WriteCell(string filePath, string sheetName, string cellAddress, object value)
        {
            // --- Cél: Egyszerűség és újrafelhasználás ---
            // Ez a metódus egy "parancsikon". Ahelyett, hogy újraimplementálnánk a fájlkezelést,
            // létrehoz egy ExcelManager példányt egy 'using' blokkban. Ez garantálja, hogy
            // ugyanaz a validálási és mentési logika fut le, mint a komplexebb használat esetén.
            using (var manager = new ExcelManager(filePath))
            {
                var worksheet = manager.Worksheet(sheetName);
                worksheet.Cell(cellAddress).Value = XLCellValue.FromObject(value);
            }
        }
    }
}