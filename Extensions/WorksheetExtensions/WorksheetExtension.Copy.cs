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
        /// Lemásolja a munkalapot a munkafüzeten belül egy új névre.
        /// </summary>
        /// <param name="worksheet">A másolandó munkalap.</param>
        /// <param name="newSheetName">A létrehozandó másolat neve.</param>
        /// <returns>A frissen létrehozott munkalap másolata (IXLWorksheet).</returns>
        public static IXLWorksheet Copy(this IXLWorksheet worksheet, string newSheetName)
        {
            // Cél: Egyszerű, láncolható másolási parancs.
            // A ClosedXML beépített CopyTo metódusát használjuk, ami visszaadja
            // az új munkalapot, így akár tovább is láncolhatók a műveletek.
            if (string.IsNullOrWhiteSpace(newSheetName))
            {
                throw new ArgumentException("Az új munkalap neve nem lehet üres.", nameof(newSheetName));
            }

            return worksheet.CopyTo(newSheetName);
        }
    }
}
