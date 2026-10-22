using ClosedXML.Excel;
using System.Linq;

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
        /// Visszaadja a munkalap utolsó, adatot tartalmazó sorának számát.
        /// Ha a munkalap üres, 0-t ad vissza.
        /// </summary>
        /// <param name="worksheet">A munkalap, amin a műveletet végrehajtjuk.</param>
        /// <returns>Az utolsó használt sor száma.</returns>
        public static int GetLastRow(this IXLWorksheet worksheet, string columnLetter)
        {
            var column = worksheet.Column(columnLetter);
            var lastCell = column.CellsUsed().LastOrDefault();

            return lastCell?.Address.RowNumber ?? 0;
        }
    }
}