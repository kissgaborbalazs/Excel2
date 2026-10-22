using System;

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
        /// Átnevez egy munkalapot egy Excel fájlban.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap jelenlegi neve.</param>
        /// <param name="newName">A munkalap új neve.</param>
        public static void RenameSheet(string filePath, string sheetName, string newName)
        {
            using (var manager = new ExcelManager(filePath))
            {
                manager.Worksheet(sheetName).Rename(newName);
            }
        }
    }
}