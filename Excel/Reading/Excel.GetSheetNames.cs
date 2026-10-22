using System.Collections.Generic;
using System.Linq;

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
        /// Lekérdezi egy Excel fájl összes munkalapjának nevét egyetlen, atomi művelettel.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <returns>A munkalapnevek listája (List<string>).</returns>
        public static List<string> GetSheetNames(string filePath)
        {
            // Cél: Egyszerűség és kényelem biztosítása.
            // A 'using' blokk garantálja a biztonságos fájlkezelést.
            using (var manager = new ExcelManager(filePath))
            {
                // A tényleges munkát az ExcelManager-ben definiált metódusra delegálja.
                // A .ToList() fontos, mert ez "materializálja" az eredményt (végrehajtja a lekérdezést),
                // mielőtt a 'using' blokk végén a manager és a benne lévő munkafüzet-objektum
                // felszabadulna. Enélkül hibát kapnánk.
                return manager.GetSheetNames().ToList();
            }
        }
    }
}
