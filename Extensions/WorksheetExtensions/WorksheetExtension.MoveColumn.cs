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
        /// Áthelyez egy oszlopot a munkalapon belül egy új pozícióba.
        /// </summary>
        /// <param name="worksheet">A munkalap, amin a műveletet végrehajtjuk.</param>
        /// <param name="sourceColumnNumber">A mozgatandó oszlop sorszáma.</param>
        /// <param name="destinationColumnNumber">A célpozíció sorszáma.</param>
        public static void MoveColumn(this IXLWorksheet worksheet, int sourceColumnNumber, int destinationColumnNumber)
        {
            // Cél: Biztonságos és robusztus oszlopáthelyezés.
            if (worksheet is null)
                throw new ArgumentNullException(nameof(worksheet),"A worksheet paramáter nem lehet null!");
            // Ellenőrizzük a bemeneti paramétereket, hogy elkerüljük a hibákat.
            if (sourceColumnNumber <= 0 || destinationColumnNumber <= 0 || sourceColumnNumber == destinationColumnNumber)
            {
                return; // Nincs értelme a műveletnek, ha a pozíciók érvénytelenek vagy megegyeznek.
            }

            var sourceColumn = worksheet.Column(sourceColumnNumber);

            // Létrehozunk egy új, üres oszlopot a célpozícióban.
            if (destinationColumnNumber < sourceColumnNumber)
            {
                // Balra mozgatás: a cél ELÉ illesztünk, a forrás 1-gyel jobbra csúszik
                worksheet.Column(destinationColumnNumber).InsertColumnsBefore(1);
                // A forrás oszlop az Insert miatt 1-gyel jobbra tolódott

                worksheet.Column(sourceColumnNumber + 1).CopyTo(worksheet.Column(destinationColumnNumber));
                worksheet.Column(sourceColumnNumber + 1).Delete();
            }
            else // destinationColumnNumber > sourceColumnNumber
            {
                // Jobbra mozgatás: a cél MÖGÉ illesztünk, mert a későbbi törlés balra húz 1-et
                worksheet.Column(destinationColumnNumber).InsertColumnsAfter(1);
                worksheet.Column(sourceColumnNumber).CopyTo(worksheet.Column(destinationColumnNumber + 1));
                worksheet.Column(sourceColumnNumber).Delete();
            }
        }
    }
}