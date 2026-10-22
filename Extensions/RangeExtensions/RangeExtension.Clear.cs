using ClosedXML.Excel;

namespace MPRT.Excel
{
    /// <summary>
    /// Bővítő metódusokat (extension methods) tartalmaz a ClosedXML IXLWorksheet objektumokhoz,
    /// hogy a használatuk még intuitívabb és láncolhatóbb (fluent) legyen.
    /// Az itt definiált metódusok úgy viselkednek, mintha a ClosedXML beépített részei lennének.
    /// </summary>
    public static partial class RangeExtensions
    {
        /// <summary>
        /// Törli egy adott tartomány tartalmát.
        /// </summary>
        /// <param name="range">A törlendő tartomány.</param>
        /// <param name="clearStyles">Igaz, ha a tartalom mellett a formázást is törölni kell. Hamis (alapértelmezett), ha csak a tartalmat.</param>
        public static void Clear(this IXLRange range, bool clearStyles = false)
        {
//            if (clearStyles)
//            {
//                range.Clear(XLClearOptions.All);
//            }
//            else
//            {
//                // JAVÍTÁS: Cellánkénti törléssel a stílusok megmaradnak.
//                foreach (var cell in range.Cells())
//                {
//                    cell.Clear(XLClearOptions.Contents);
//                }
//            }
            
            
            var clearOptions = clearStyles ? XLClearOptions.All : XLClearOptions.Contents;
            range.Clear(clearOptions);
        }
    }
}