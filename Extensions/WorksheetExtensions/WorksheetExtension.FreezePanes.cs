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
        // ... A meglévő metódusok itt vannak ...

        /// <summary>
        /// Rögzíti a munkalap paneljeit a megadott sor és oszlop alapján.
        /// </summary>
        /// <param name="worksheet">A munkalap, amin a műveletet végrehajtjuk.</param>
        /// <param name="rowsToFreeze">A felülről rögzítendő sorok száma.</param>
        /// <param name="columnsToFreeze">A balról rögzítendő oszlopok száma.</param>
        public static void FreezePanes(this IXLWorksheet worksheet, int rowsToFreeze, int columnsToFreeze)
        {
            // A panelek rögzítése alapvető elvárás a nagy, görgethető táblázatoknál.
            // A ClosedXML beépített Freeze() metódusát használjuk a tiszta implementációért.
            if (rowsToFreeze >= 0 && columnsToFreeze >= 0)
            {
                worksheet.SheetView.Freeze(rowsToFreeze, columnsToFreeze);
            }
        }
    }
}
