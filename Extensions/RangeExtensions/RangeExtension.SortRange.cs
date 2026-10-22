using System;
using System.Linq;
using ClosedXML.Excel;

namespace MPRT.Excel
{
    public static partial class RangeExtensions
    {
        /// <summary>
        /// Rendezi a tartományt a megadott oszlop sorszáma alapján.
        /// </summary>
        /// <param name="range">A rendezendő tartomány.</param>
        /// <param name="columnNumber">A rendezés alapjául szolgáló oszlop sorszáma a tartományon belül (1-től kezdődően).</param>
        /// <param name="sortOrder">A rendezés iránya (növekvő vagy csökkenő).</param>
        /// <param name="hasHeader">Igaz, ha a tartomány első sorát fejlécként kell kezelni (nem vonja be a rendezésbe).</param>
        public static void SortRange(this IXLRange range, int columnNumber, XLSortOrder sortOrder = XLSortOrder.Ascending, bool hasHeader = true)
        {
            // Cél: Egyszerű, parancs-szerű rendezés biztosítása.
            if (columnNumber > range.ColumnCount())
                throw new ArgumentOutOfRangeException(nameof(columnNumber), $"A megadott oszlop szám ({columnNumber.ToString()}, kívűl esik az aktuális tartományon)");

            // Ha van fejléc, akkor az első sort kihagyjuk a rendezésből
            var rangeToSort = hasHeader
                ? range.Range(2, 1, range.RowCount(), range.ColumnCount())
                : range;
            Console.WriteLine(rangeToSort.ToString());
            rangeToSort.Sort(columnNumber, sortOrder);
        }

        /// <summary>
        /// Rendezi a tartományt a megadott oszlop betűjele alapján.
        /// </summary>
        /// <param name="range">A rendezendő tartomány.</param>
        /// <param name="columnLetter">A rendezés alapjául szolgáló oszlop betűjele (pl. "A", "C").</param>
        /// <param name="sortOrder">A rendezés iránya (növekvő vagy csökkenő).</param>
        /// <param name="hasHeader">Igaz, ha a tartomány első sorát fejlécként kell kezelni (nem vonja be a rendezésbe).</param>
        public static void SortRange(this IXLRange range, string columnLetter, XLSortOrder sortOrder = XLSortOrder.Ascending, bool hasHeader = true)
        {
            // Cél: Egyszerű, parancs-szerű rendezés biztosítása.
            if (string.IsNullOrWhiteSpace(columnLetter))
                throw new ArgumentNullException(nameof(columnLetter), "Az oszlop azonosítója nem lehet null!");

            int columnNumber = range.Column(columnLetter).ColumnNumber() - range.FirstCell().Address.ColumnNumber + 1;

            if (columnNumber < 1 || columnNumber > range.ColumnCount())
                throw new ArgumentOutOfRangeException(nameof(columnNumber), $"A megadott oszlop ({columnNumber.ToString()}, kívül esik az aktuális tartományon)");

            range.SortRange(columnNumber, sortOrder, hasHeader);
        }

        /// <summary>
        /// Rendezi a tartományt a megadott oszlop neve alapján.
        /// A metódus a tartomány első sorában (fejléc) keresi a megadott oszlopnevet.
        /// </summary>
        /// <param name="range">A rendezendő tartomány.</param>
        /// <param name="columnName">A rendezés alapjául szolgáló oszlop neve.</param>
        /// <param name="sortOrder">A rendezés iránya (növekvő vagy csökkenő).</param>
        public static void SortRangeByName(this IXLRange range, string columnName, XLSortOrder sortOrder = XLSortOrder.Ascending)
        {
            // Cél: Felhasználóbarát, név alapú rendezés lehetővé tétele.
            if (string.IsNullOrWhiteSpace(columnName))
                throw new ArgumentNullException(nameof(columnName), "Az oszlopnév nem lehet üres!");

            // 1. Megkeressük az oszlop sorszámát a fejlécben.
            // Végigmegyünk a tartomány első sorának celláin.
            var headerRow = range.FirstRow();
            var columnCell = headerRow.Cells()
                .FirstOrDefault(c => c.GetText().Equals(columnName, StringComparison.OrdinalIgnoreCase));

            if (columnCell == null)
            {
                // Ha nem találjuk az oszlopot, informatív hibát dobunk.
                throw new ArgumentException($"A '{columnName}' nevű oszlop nem található a tartomány fejlécében.", nameof(columnName));
            }

            // 2. Meghatározzuk a relatív oszlopszámot a tartományon belül.
            // A ColumnNumber() a teljes munkalaphoz viszonyított sorszámot adja,
            // nekünk a tartomány kezdőoszlopához viszonyított kell.
            int sortColumnIndex = columnCell.Address.ColumnNumber - range.FirstCell().Address.ColumnNumber + 1;

            // 3. Meghívjuk a sorszám alapú rendezést.
            range.SortRange(sortColumnIndex, sortOrder, true);
        }
    }
}