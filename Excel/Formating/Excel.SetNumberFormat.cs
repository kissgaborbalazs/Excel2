using ClosedXML.Excel;

namespace MPRT.Excel
{
    public static partial class Excel
    {
        /// <summary>
        /// Beállítja egy megadott tartomány számformátumát egy Excel fájlban egyedi string alapján.
        /// </summary>
        public static void SetNumberFormat(string filePath, string sheetName, string rangeAddress, string format)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(sheetName);
                var range = worksheet.Range(rangeAddress);
                range.SetNumberFormat(format);
                workbook.Save();
            }
        }

        /// <summary>
        /// Beállítja egy megadott tartomány számformátumát egy Excel fájlban előre definiált típus alapján.
        /// </summary>
        public static void SetNumberFormat(string filePath, string sheetName, string rangeAddress, NumberFormatType formatType)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(sheetName);
                var range = worksheet.Range(rangeAddress);
                range.SetNumberFormat(formatType);
                workbook.Save();
            }
        }
    }
}