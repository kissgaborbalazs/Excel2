using ClosedXML.Excel;
using System;
using System.Linq;

namespace MPRT.Excel
{
    public static partial class WorksheetExtensions
    {
        /// <summary>
        /// Megkeresi az első cellát, amely a megadott értéket tartalmazza.
        /// A keresés nem tesz különbséget a kis- és nagybetűk között.
        /// </summary>
        /// <param name="worksheet">A munkalap, amin a keresést végrehajtjuk.</param>
        /// <param name="valueToFind">A keresendő érték.</param>
        /// <returns>A megtalált cella címe (pl. "A5"), vagy üres string, ha nem található.</returns>
        /// <remarks>
        /// ⚠️ Ismert korlát: A tömbképleteket (array formula, pl. <c>{=SUM(...)}</c>) tartalmazó cellák
        /// nem kereshetők — a ClosedXML nem támogatja ezek kiértékelését.
        /// Ilyen cellák a keresés során kihagyásra kerülnek.
        /// </remarks>
        public static string FindCell(this IXLWorksheet worksheet, object valueToFind)
        {
            if (valueToFind == null) return string.Empty;

            string searchValue = valueToFind.ToString();

            var cell = worksheet.CellsUsed()
                                .FirstOrDefault(c => SafeGetCellValue(c).Equals(searchValue, StringComparison.OrdinalIgnoreCase));

            return cell?.Address.ToString() ?? string.Empty;
        }

        /// <summary>
        /// Biztonságosan olvassa ki a cella értékét stringként.
        /// Tömbképletes celláknál a ClosedXML NotImplementedException-t dob —
        /// ezt elkapja és üres stringgel tér vissza.
        /// </summary>
        private static string SafeGetCellValue(IXLCell cell)
        {
            try { return cell.Value.ToString(); }
            catch (NotImplementedException) { return string.Empty; }
        }
    }
}