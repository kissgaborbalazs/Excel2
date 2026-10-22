using ClosedXML.Excel;
using System;
using System.Data;
using System.Linq;

namespace MPRT.Excel
{
    /// <summary>
    /// Bővítő metódusokat (extension methods) tartalmaz a ClosedXML IXLWorksheet objektumokhoz.
    /// </summary>
    public static partial class WorksheetExtensions
    {
        /// <summary>
        /// Intelligensen meghatároz egy tartományt a munkalapon.
        /// Ha a kezdőcella nincs megadva, a teljes használt területet adja vissza.
        /// Ha a kezdőcella meg van adva, a tőle jobbra és lefelé terjeszkedő, összefüggő táblázatot keresi meg.
        /// </summary>
        /// <param name="worksheet">A munkalap, amin a műveletet végrehajtjuk.</param>
        /// <param name="startCell">A tartomány kezdőcellája. Ha nincs megadva, a teljes használt területet adja vissza.</param>
        /// <returns>A megtalált tartomány IXLRange objektumként.</returns>
        public static IXLRange GetRange(this IXLWorksheet worksheet, string startCell)
        {
            if (worksheet is null)
                throw new ArgumentNullException(nameof(worksheet));

            // Ha nincs kezdőcella megadva, a teljes használt területet adjuk vissza.
            if (string.IsNullOrWhiteSpace(startCell))
            {
                return worksheet.RangeUsed();
            }

            var firstCell = worksheet.Cell(startCell);
            int firstCol = firstCell.Address.ColumnNumber;
            int firstRow = firstCell.Address.RowNumber;

            // 1. A táblázat szélességének meghatározása a kezdő sor alapján.
            // Végigmegyünk a kezdő sorban lévő, nem üres cellákon.
            var headerCells = worksheet.Row(firstRow)
                .CellsUsed(c => c.Address.ColumnNumber >= firstCol);
            
            // Ha a kezdő sor üres, akkor csak az egyetlen kezdőcellát adjuk vissza.
            if (!headerCells.Any())
            {
                return firstCell.AsRange();
            }

            int lastCol = headerCells.Last().Address.ColumnNumber;

            // 2. A táblázat magasságának meghatározása.
            // Végigmegyünk a táblázat minden oszlopán (a szélesség alapján),
            // és megkeressük bennük az utolsó, adatot tartalmazó sort.
            int lastRow = headerCells
                .Select(headerCell =>
                    worksheet.Column(headerCell.Address.ColumnNumber)
                             .CellsUsed(c => c.Address.RowNumber >= firstRow)
                             .LastOrDefault()?.Address.RowNumber ?? 0)
                .Max();
            
            // Ha valamiért a lastRow kisebb lenne, mint a kezdő sor, korrigálunk.
            if (lastRow < firstRow) lastRow = firstRow;

            // 3. A végső tartomány létrehozása a kiszámított határok alapján.
            return worksheet.Range(firstRow, firstCol, lastRow, lastCol);
        }
    }
}