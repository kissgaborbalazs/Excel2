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
        /// Törli a munkalap teljes tartalmát.
        /// </summary>
        /// <param name="worksheet">A törlendő munkalap.</param>
        /// <param name="clearStyles">Igaz, ha a tartalom mellett a formázást is törölni kell. Hamis (alapértelmezett), ha csak a tartalmat.</param>
        public static void Clear(this IXLWorksheet worksheet, bool clearStyles = false)
        {
            // Cél: Egységes törlési logika biztosítása.
            // A clearOptions változóval eldöntjük, hogy milyen mélységű törlést végzünk.
            // A ClosedXML beépített Clear() metódusa elvégzi a tényleges munkát.
            var clearOptions = clearStyles ? XLClearOptions.All : XLClearOptions.Contents;
            worksheet.Clear(clearOptions);
        }
    }
}