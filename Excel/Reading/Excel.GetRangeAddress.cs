namespace MPRT.Excel
{
    /// <summary>
    /// Egyszerű, statikus segédmetódusokat biztosít a leggyakoribb Excel műveletekhez.
    /// (Ez a class partial, a metódusok külön fájlokban vannak a jobb szervezettség érdekében.)
    /// </summary>
    public static partial class Excel
    {
        /// <summary>
        /// Intelligensen meghatározza egy tartomány címét a megadott kezdőcella alapján.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap neve.</param>
        /// <param name="startCell">A tartomány kezdőcellája. Ha nincs megadva, a teljes használt terület címét adja vissza.</param>
        /// <returns>A megtalált tartomány címe (pl. "A2:D15").</returns>
        public static string GetRangeAddress(string filePath, string sheetName, string startCell = null)
        {
            // Cél: Egy kényelmi funkció biztosítása, ami a dinamikus tartománykeresést
            // egy egyszerű, stringet visszaadó metódussá alakítja.
            using (var manager = new ExcelManager(filePath))
            {
                // A tényleges logikát a már meglévő, robusztus GetRange bővítő metódusra bízzuk.
                var range = manager.Worksheet(sheetName).GetRange(startCell);

                // A megtalált tartomány címét adjuk vissza. Ha a tartomány üres,
                // a .ToString() üres stringet ad, ami biztonságosan kezelhető.
                return range.RangeAddress.ToString();
            }
        }
    }
}