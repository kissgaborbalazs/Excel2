using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace MPRT.Excel
{
    /// <summary>
    /// Bővítő metódusokat (extension methods) tartalmaz a ClosedXML IXLWorkbook objektumokhoz,
    /// hogy a használatuk még intuitívabb és láncolhatóbb (fluent) legyen.
    /// Az itt definiált metódusok úgy viselkednek, mintha a ClosedXML beépített részei lennének.
    /// </summary>
    public static partial class WorkbookExtensions
    {
        /// <summary>
        /// Ellenőrzi, hogy egy adott nevű munkalap létezik-e a munkafüzetben.
        /// A keresés nem tesz különbséget a kis- és nagybetűk között.
        /// </summary>
        /// <param name="workbook">A munkafüzet, amin a műveletet végrehajtjuk.</param>
        /// <param name="sheetName">A keresendő munkalap neve.</param>
        /// <returns>Igaz (true), ha a munkalap létezik; egyébként hamis (false).</returns>
        public static bool SheetExists(this IXLWorkbook workbook, string sheetName)
        {
            // Ellenőrzés a metódus elején, hogy elkerüljük a kivételt.
            if (string.IsNullOrWhiteSpace(sheetName))
                return false;

            // Cél: Gyors és megbízható ellenőrzés.
            // A ClosedXML beépített 'Contains' metódusa a leghatékonyabb módja
            // a munkalap létezésének ellenőrzésére.
            return workbook.Worksheets.Contains(sheetName);
        }
    }
}
