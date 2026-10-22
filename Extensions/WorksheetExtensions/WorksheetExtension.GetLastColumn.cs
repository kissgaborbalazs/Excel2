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
        /// Visszaadja az adott sor utolsó, használt oszlopának számát.
        /// Ha a munkalap üres, 0-t ad vissza.
        /// </summary>
        /// <param name="worksheet">A munkalap, amin a műveletet végrehajtjuk.</param>
        /// <returns>Az utolsó használt oszlop sorszáma.</returns>
        public static string GetLastColumn(this IXLWorksheet worksheet, int rowNumber)
        {
            var row = worksheet.Row(rowNumber);
            var lastCell = row.CellsUsed().LastOrDefault();
            return lastCell?.Address.ColumnLetter ?? string.Empty;
        }
    }
}
