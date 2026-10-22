using ClosedXML.Excel;
using System.Data;

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
        /// Hozzáfűz egy DataTable tartalmát egy munkalap meglévő adatai után.
        /// A művelet a fejléceket figyelmen kívül hagyja, csak az adatokat írja be.
        /// </summary>
        /// <param name="worksheet">A munkalap, amihez hozzá szeretnénk fűzni az adatokat.</param>
        /// <param name="dataTable">A hozzáfűzendő adatok.</param>
        public static void AppendDataTable(this IXLWorksheet worksheet, DataTable dataTable)
        {
            // Cél: Elkerülni a felesleges írást és a hibákat, ha nincs mit hozzáfűzni.
            if (dataTable == null || dataTable.Rows.Count == 0) return;

            // Cél: A kezdő cella dinamikus meghatározása.
            var lastRow = worksheet.LastRowUsed();
            int startRow = lastRow is null ? 1: lastRow.RowNumber() + 1;
            string startCell = $"A{startRow}";

            // Cél: Kód-újrahasznosítás.
            // A már meglévő, jól bevált WriteDataTable metódust hívjuk meg,
            // de explicit módon jelezzük neki, hogy a fejléceket ne írja be.
            worksheet.WriteDataTable(dataTable, startCell, false);
        }
    }
}