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
        /// Visszaadja a munkafüzetben található összes munkalap nevét.
        /// </summary>
        /// <param name="workbook">A munkafüzet, amin a műveletet végrehajtjuk.</param>
        /// <returns>A munkalapnevek listája (IEnumerable<string>).</returns>
        public static IEnumerable<string> GetSheetNames(this IXLWorkbook workbook)
        {
            // Cél: Gyors és hatékony munkalapnév-lekérdezés.
            // A LINQ Select metódusával egyszerűen kivetítjük a munkalap-kollekció
            // minden elemének (ws) a nevét (Name) egy új kollekcióba.
            return workbook.Worksheets.Select(ws => ws.Name);
        }
    }
}
