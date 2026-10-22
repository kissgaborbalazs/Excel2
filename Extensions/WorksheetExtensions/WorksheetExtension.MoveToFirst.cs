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
        /// A munkalapot az első pozícióba helyezi.
        /// </summary>
        /// <param name="worksheet">A mozgatandó munkalap.</param>
        public static void MoveSheetToFirst(this IXLWorksheet worksheet)
        {
            // Cél: Kényelmi funkció a leggyakoribb pozicionálási feladatra.
            if (worksheet is null)
                throw new ArgumentNullException(nameof(worksheet));
            // A ClosedXML-ben a pozíció 1-től indexelődik.
            worksheet.Position = 1;
        }
    }
}