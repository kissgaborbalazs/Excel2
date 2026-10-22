namespace MPRT.Excel
{
    /// <summary>
    /// Egyszerű, statikus segédmetódusokat biztosít a leggyakoribb Excel műveletekhez.
    /// (Ez a class partial, a metódusok külön fájlokban vannak a jobb szervezettség érdekében.)
    /// </summary>
    public static partial class Excel
    {
        /// <summary>
        /// Beállítja egy megadott tartomány stílusát egy Excel fájlban egy Style objektum alapján.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap neve.</param>
        /// <param name="rangeAddress">A formázandó tartomány címe (pl. "A1:D1").</param>
        /// <param name="style">A stílusbeállításokat tartalmazó objektum.</param>
        public static void SetRangeStyle(string filePath, string sheetName, string rangeAddress, Style style)
        {
            // Cél: Egyszerűség és kényelem biztosítása a komplex stílusok beállításához.
            // A 'using' blokk garantálja a biztonságos fájlkezelést.
            using (var manager = new ExcelManager(filePath))
            {
                // A tényleges munkát a RangeExtensions-ben definiált, láncolható SetStyle metódusra delegálja.
                manager.Worksheet(sheetName).Range(rangeAddress).SetStyle(style);
            }
        }
    }
}