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
        /// Törli a munkalapot a munkafüzetből.
        /// Figyelem: A művelet után a worksheet objektum érvénytelenné válik.
        /// </summary>
        /// <param name="worksheet">A törlendő munkalap.</param>
        public static void Delete(this IXLWorksheet worksheet)
        {
            // Cél: Egyszerű, láncolható törlési parancs biztosítása.
            // A ClosedXML beépített Delete() metódusát hívjuk meg.
            worksheet.Delete();
        }
    }
}