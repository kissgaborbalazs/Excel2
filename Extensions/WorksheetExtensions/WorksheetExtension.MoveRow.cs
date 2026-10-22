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
        /// Áthelyez egy sort a munkalapon belül egy új pozícióba.
        /// </summary>
        /// <param name="worksheet">A munkalap, amin a műveletet végrehajtjuk.</param>
        /// <param name="sourceRowNumber">A mozgatandó sor sorszáma.</param>
        /// <param name="destinationRowNumber">A célpozíció sorszáma.</param>
        public static void MoveRow(this IXLWorksheet worksheet, int sourceRowNumber, int destinationRowNumber)
        {
            // Cél: Biztonságos és robusztus soráthelyezés.
            // Ellenőrizzük a bemeneti paramétereket, hogy elkerüljük a hibákat.
            if (worksheet is null)
                throw new ArgumentNullException(nameof(worksheet));

            if (sourceRowNumber <= 0 || destinationRowNumber <= 0 || sourceRowNumber == destinationRowNumber)
            {
                return; // Nincs értelme a műveletnek, ha a pozíciók érvénytelenek vagy megegyeznek.
            }

            if (destinationRowNumber < sourceRowNumber)
            {
                // Felfelé mozgatás: cél SOR fölé szúrunk
                worksheet.Row(destinationRowNumber).InsertRowsAbove(1);
                worksheet.Row(sourceRowNumber + 1).CopyTo(worksheet.Row(destinationRowNumber));
                worksheet.Row(sourceRowNumber + 1).Delete();
            }
            else // destinationRowNumber > sourceRowNumber
            {
                // Lefelé mozgatás: cél SOR alá szúrunk, mert a későbbi törlés 1-gyel felfelé húz
                worksheet.Row(destinationRowNumber).InsertRowsBelow(1);
                worksheet.Row(sourceRowNumber).CopyTo(worksheet.Row(destinationRowNumber + 1));
                worksheet.Row(sourceRowNumber).Delete();
            }
        }
    }
}