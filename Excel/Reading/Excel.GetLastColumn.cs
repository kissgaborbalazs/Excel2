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
        /// Visszaadja egy adott sor utolsó, adatot tartalmazó oszlopának betűjelét.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap neve.</param>
        /// <param name="rowNumber">A sor sorszáma, amiben keresünk.</param>
        /// <returns>Az utolsó használt oszlop betűjele (pl. "D"), vagy üres string, ha a sor üres.</returns>
        public static string GetLastColumn(string filePath, string sheetName, int rowNumber)
        {
            using (var manager = new ExcelManager(filePath))
            {
                // Itt a te új, módosított metódusodat hívnánk meg.
                // A példa kedvéért feltételezem, hogy a metódus neve a javaslat szerint módosult.
                return manager.Worksheet(sheetName).GetLastColumn(rowNumber);
            }
        }
    }
}
