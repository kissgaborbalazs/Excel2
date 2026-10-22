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
        /// Egy munkalapot az utolsó pozícióba helyez egy Excel fájlban.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A mozgatandó munkalap neve.</param>
        public static void MoveSheetToLast(string filePath, string sheetName)
        {
            using (var manager = new ExcelManager(filePath))
            {
                manager.Worksheet(sheetName).MoveSheetToLast();
            }
        }
    }
}