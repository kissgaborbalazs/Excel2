using ClosedXML.Excel;

namespace MPRT.Excel
{
    /// <summary>
    /// Egyszerű, statikus segédmetódusokat biztosít a leggyakoribb Excel műveletekhez.
    /// Minden metódus egy teljes, önálló műveletet hajt végre (megnyitás, módosítás, mentés, bezárás).
    /// Akkor ideális, ha csak egyetlen műveletet kell elvégezni egy fájlon.
    /// </summary>
    public static partial class Excel
    {
        /// <summary>
        /// Rendezi a megadott tartományt egy Excel fájlban a tartományon belüli oszlopindex alapján.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap neve, amely a tartományt tartalmazza.</param>
        /// <param name="rangeAddress">A rendezendő tartomány címe (pl. "A1:D10").</param>
        /// <param name="columnNumber">A rendezés alapjául szolgáló oszlop sorszáma a megadott tartományon belül (1-től kezdődően).</param>
        /// <param name="sortOrder">A rendezés iránya (növekvő vagy csökkenő). Alapértelmezett a növekvő.</param>
        /// <param name="hasHeader">Meghatározza, hogy a tartomány tartalmaz-e fejléc sort. Ha <c>true</c>, a fejléc nem kerül bele a rendezésbe.</param>
        public static void SortRange(string filePath, string sheetName, string rangeAddress, int columnNumber, XLSortOrder sortOrder = XLSortOrder.Ascending, bool hasHeader = true)
        {
            using (var manager = new ExcelManager(filePath))
            {
                manager.Worksheet(sheetName).Range(rangeAddress).SortRange(columnNumber, sortOrder, hasHeader);
            }
        }

        /// <summary>
        /// Rendezi a megadott tartományt egy Excel fájlban az oszlop betűjele alapján.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap neve, amely a tartományt tartalmazza.</param>
        /// <param name="rangeAddress">A rendezendő tartomány címe (pl. "A1:D10").</param>
        /// <param name="columnLetter">A rendezés alapjául szolgáló oszlop betűjele (pl. "A", "B").</param>
        /// <param name="sortOrder">A rendezés iránya (növekvő vagy csökkenő). Alapértelmezett a növekvő.</param>
        /// <param name="hasHeader">Meghatározza, hogy a tartomány tartalmaz-e fejléc sort. Ha <c>true</c>, a fejléc nem kerül bele a rendezésbe.</param>
        public static void SortRange(string filePath, string sheetName, string rangeAddress, string columnLetter, XLSortOrder sortOrder = XLSortOrder.Ascending, bool hasHeader = true)
        {
            using (var manager = new ExcelManager(filePath))
            {
                // Meghívja a RangeExtensions-ben definiált, név alapú Sort metódust.
                manager.Worksheet(sheetName).Range(rangeAddress).SortRange(columnLetter, sortOrder, hasHeader);
            }
        }
/// <summary>
        /// Rendezi a megadott tartományt egy Excel fájlban a fejlécben szereplő oszlopnév alapján.
        /// A rendezéshez a tartománynak rendelkeznie kell fejléccel.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <param name="sheetName">A munkalap neve, amely a tartományt tartalmazza.</param>
        /// <param name="rangeAddress">A rendezendő tartomány címe (pl. "A1:D10").</param>
        /// <param name="columnName">A rendezés alapjául szolgáló oszlop neve a fejléc sorban.</param>
        /// <param name="sortOrder">A rendezés iránya (növekvő vagy csökkenő). Alapértelmezett a növekvő.</param>
        /// <param name="hasHeader">Meghatározza, hogy a tartomány tartalmaz-e fejléc sort. A név alapú rendezéshez ennek <c>true</c> értékűnek kell lennie.</param>
        public static void SortRangeByName(string filePath, string sheetName, string rangeAddress, string columnName, XLSortOrder sortOrder = XLSortOrder.Ascending, bool hasHeader = true)
        {
            using (var manager = new ExcelManager(filePath))
            {
                manager.Worksheet(sheetName).Range(rangeAddress).SortRangeByName(columnName, sortOrder);
            }
        }
    }
}