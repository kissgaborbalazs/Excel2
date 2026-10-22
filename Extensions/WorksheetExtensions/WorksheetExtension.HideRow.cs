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
        /// Elrejt egy sort a munkalapon.
        /// </summary>
        /// <param name="worksheet">A munkalap, amin a műveletet végrehajtjuk.</param>
        /// <param name="rowNumber">Az elrejtendő sor sorszáma.</param>
        public static void HideRow(this IXLWorksheet worksheet, int rowNumber)
        {
            // Cél: Egyszerű, parancs-szerű sor elrejtés.
            // A ClosedXML beépített Hide() metódusát hívjuk meg a megfelelő sor objektumon.
            if (rowNumber > 0)
            {
                worksheet.Row(rowNumber).Hide();
            }
        }
    }
}
