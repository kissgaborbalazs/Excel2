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
        /// Visszaadja egy adott oszlop utolsó, adatot tartalmazó sorának sorszámát.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap neve.</param>
        /// <param name="columnLetter">Annak az oszlopnak a betűjele, amiben keresünk (pl. "A").</param>
        /// <returns>Az utolsó használt sor sorszáma az adott oszlopban.</returns>
        public static int GetLastRow(string filePath, string sheetName, string columnLetter)
        {
            using (var manager = new ExcelManager(filePath))
            {
                // A te új, módosított metódusodat hívja meg.
                return manager.Worksheet(sheetName).GetLastRow(columnLetter);
            }
        }
    }
}
