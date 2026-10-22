using ClosedXML.Excel;

namespace MPRT.Excel
{
    /// <summary>
    /// Bővítő metódusokat (extension methods) tartalmaz a ClosedXML IXLWorksheet objektumokhoz,
    /// hogy a használatuk még intuitívabb és láncolhatóbb (fluent) legyen.
    /// Az itt definiált metódusok úgy viselkednek, mintha a ClosedXML beépített részei lennének.
    /// </summary>
    public static partial class WorksheetExtensions
    {
        /// <summary>
        /// A megadott oszlopok szélességét a tartalmukhoz igazítja.
        /// Ha nincsenek oszlopok megadva, az összes oszlopot igazítja.
        /// </summary>
        /// <param name="worksheet">A munkalap, amin a műveletet végrehajtjuk.</param>
        /// <param name="columns">Az igazítandó oszlopok sorszámai (pl. 1, 3, 5).</param>
        public static void AutoFitColumns(this IXLWorksheet worksheet, params int[] columns)
        {
            // Cél: Rugalmas oszlopigazítás biztosítása.
            // A 'params' kulcsszó lehetővé teszi, hogy a felhasználó tetszőleges számú
            // oszlopszámot adjon meg, vagy egyet sem.

            if (columns == null || columns.Length == 0)
            {
                // Ha nincs konkrét oszlop megadva, az összes oszlopot igazítjuk.
                // A Columns() metódus a munkalap összes oszlopát visszaadja.
                worksheet.Columns().AdjustToContents();
            }
            else
            {
                // Ha vannak megadott oszlopok, egyenként végigmegyünk rajtuk
                // és csak azokat igazítjuk.
                foreach (var colNum in columns)
                {
                    if (colNum > 0) // Biztonsági ellenőrzés az érvényes oszlopszámra
                    {
                        worksheet.Column(colNum).AdjustToContents();
                    }
                }
            }
        }
    }
}