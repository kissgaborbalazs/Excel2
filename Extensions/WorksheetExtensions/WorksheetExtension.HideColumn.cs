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
        /// Elrejt egy oszlopot a munkalapon.
        /// </summary>
        /// <param name="worksheet">A munkalap, amin a műveletet végrehajtjuk.</param>
        /// <param name="columnNumber">Az elrejtendő oszlop sorszáma.</param>
        public static void HideColumn(this IXLWorksheet worksheet, int columnNumber)
        {
            // Cél: Egyszerű, parancs-szerű oszlop elrejtés.
            // A ClosedXML beépített Hide() metódusát hívjuk meg a megfelelő oszlop objektumon.
            if (columnNumber > 0)
            {
                worksheet.Column(columnNumber).Hide();
            }
        }
    }
}