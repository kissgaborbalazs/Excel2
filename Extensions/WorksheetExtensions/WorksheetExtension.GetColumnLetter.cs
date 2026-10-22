using ClosedXML.Excel;
using System;
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
        /// Visszaadja a munkalapon található megadott oszlopnevet tartalmazó oszlop betűjét.
        /// Ha a munkalap üres, null-t ad vissza.
        /// </summary>
        /// <param name="worksheet">A munkalap, amin a műveletet végrehajtjuk.</param>
        /// <param name="columnName">Az oszlop neve, aminek megszeretnénk találni az oszlop betűjét.</param>
        /// <returns>Az utolsó használt sor száma.</returns>
        public static string GetColumnLetter(this IXLWorksheet worksheet, string columnName, int headerRow)
        {
            return worksheet.Row(headerRow).Cells().FirstOrDefault(c => c.GetString().Equals(columnName, StringComparison.OrdinalIgnoreCase))?.Address.ColumnLetter;
        }
    }
}