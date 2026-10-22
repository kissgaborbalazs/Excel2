using System.Data;

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
        /// Hozzáfűz egy DataTable tartalmát egy munkalap meglévő adatai után egyetlen, atomi művelettel.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap neve.</param>
        /// <param name="dataTable">A hozzáfűzendő adatok.</param>
        public static void AppendDataTable(string filePath, string sheetName, DataTable dataTable)
        {
            // Cél: Egyszerűség és kényelem biztosítása.
            // Ez a metódus elrejti a fájlkezelés komplexitását. A 'using' blokk garantálja,
            // hogy a fájl mindig biztonságosan bezárásra kerül a művelet után.
            using (var manager = new ExcelManager(filePath))
            {
                // A tényleges munkát a 'WorksheetExtensions'-ben definiált, láncolható AppendDataTable metódusra delegálja.
                manager.Worksheet(sheetName).AppendDataTable(dataTable);
            }
        }
    }
}