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
        /// Beolvas egyetlen cella értékét egy atomi művelettel.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap neve.</param>
        /// <param name="cellAddress">A beolvasandó cella címe (pl. "A1").</param>
        /// <param name="keepFormatting">Igaz, ha a cella formázott, látható értékét (string) kell beolvasni. Hamis (alapértelmezett), ha a nyers, típusos adatot (double, DateTime, stb.).</param>
        /// <returns>A cella értéke 'object' típusként. A tényleges típus a cella tartalmától és a 'keepFormatting' paramétertől függ.</returns>
        public static object ReadCell(string filePath, string sheetName, string cellAddress, bool keepFormatting = false)
        {
            // Cél: Egyszerűség és kényelem biztosítása egyetlen cella olvasásához.
            // A metódus elrejti a fájlkezelés komplexitását. A 'using' blokk garantálja,
            // hogy a fájl mindig biztonságosan bezárásra kerül a művelet után.
            using (var manager = new ExcelManager(filePath))
            {
                // Lekérjük a munkalapot és a cellát.
                var cell = manager.Worksheet(sheetName).Cell(cellAddress);

                // Cél: Választási lehetőség a nyers és a formázott érték között.
                // A 'keepFormatting' paraméter dönti el, hogy a .Value (pl. 45365.0, egy DateTime objektum)
                // vagy a .GetFormattedString() (pl. "15 000 Ft", "2024.03.14.") kerüljön visszaadásra.
                // Az alapértelmezett 'false' a típusos adatoknak kedvez, ami a számításokhoz és a
                // további feldolgozáshoz előnyösebb.
                return keepFormatting ? cell.GetFormattedString() : cell.Value;
            }
        }
    }
}
