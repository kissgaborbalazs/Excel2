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
        /// Lemásol egy munkalapot egy Excel fájlon belül.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sourceSheetName">A másolandó munkalap neve.</param>
        /// <param name="newSheetName">A létrehozandó másolat neve.</param>
        public static void CopySheet(string filePath, string sourceSheetName, string newSheetName)
        {
            using (var manager = new ExcelManager(filePath))
            {
                manager.Worksheet(sourceSheetName).Copy(newSheetName);
            }
        }
    }
}