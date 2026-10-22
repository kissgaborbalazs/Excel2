using System;

namespace MPRT.Excel
{
    public static partial class Excel
    {
        /// <summary>
        /// Külső URL hyperlinket állít be a megadott cellára. A cella látható szövege
        /// (meglévő értéke) változatlan marad, és a linkre klasszikus kék + aláhúzott stílus kerül.
        /// </summary>
        /// <param name="filePath">Az .xlsx fájl teljes elérési útja.</param>
        /// <param name="sheetName">A munkalap neve (ha nem létezik, létrejön).</param>
        /// <param name="cellAddress">A cella címe (pl. "B7").</param>
        /// <param name="url">A külső URL – abszolút, sémával (pl. "https://...").</param>
        /// <exception cref="ArgumentNullException">Ha a cellacím vagy az URL null/üres.</exception>
        /// <exception cref="ArgumentException">Ha az URL nem jól formázott abszolút URI.</exception>
        public static void SetHyperlink(string filePath, string sheetName, string cellAddress, string url)
        {
            using (var manager = new ExcelManager(filePath))
                manager.Worksheet(sheetName).SetHyperlink(cellAddress, url);
        }
    }
}