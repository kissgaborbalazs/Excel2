using System;
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
        /// Kiolvassa egy cella képletét stringként.
        /// </summary>
        /// <param name="worksheet">A munkalap, amin a műveletet végrehajtjuk.</param>
        /// <param name="cellAddress">A cella címe, aminek a képletét kiolvassuk.</param>
        /// <returns>A cella képlete stringként, vagy üres string, ha a cella nem tartalmaz képletet.</returns>
        public static string ReadFormula(this IXLWorksheet worksheet, string cellAddress)
        {
            // Cél: Lehetővé tenni a meglévő képletek ellenőrzését vagy másolását.
            if (worksheet is null)
                throw new ArgumentNullException(nameof(worksheet));

            // A HasFormula property-vel ellenőrizzük, hogy a cella valóban tartalmaz-e képletet,
            // így elkerüljük a felesleges hibákat és üres értékeket.
            if (!string.IsNullOrWhiteSpace(cellAddress))
            {
                var cell = worksheet.Cell(cellAddress);
                return cell.HasFormula ? cell.FormulaA1 : string.Empty;
            }
            return string.Empty;
        }
    }
}
