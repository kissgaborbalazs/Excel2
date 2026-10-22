using ClosedXML.Excel;
using System;

namespace MPRT.Excel
{
    /// <summary>
    /// Előre definiált, gyakran használt Excel számformátumok.
    /// </summary>
    public enum NumberFormatType
    {
        /// <summary>Formátum: "General"</summary>
        General,
        /// <summary>Formátum: "@"</summary>
        Text,
        /// <summary>Formátum: "0"</summary>
        Number,
        /// <summary>Formátum: "#,##0.00"</summary>
        NumberWithTwoDecimals,
        /// <summary>Formátum: "#,##0.00 [$Ft-40E]"</summary>
        Currency,
        /// <summary>Formátum: "yyyy.mm.dd"</summary>
        Date,
        /// <summary>Formátum: "yyyy.mm.dd hh:mm:ss"</summary>
        DateTime,
        /// <summary>Formátum: "0.00%"</summary>
        Percentage
    }

    public static partial class RangeExtensions
    {
        /// <summary>
        /// Beállítja egy adott tartomány számformátumát egy előre definiált típus alapján.
        /// </summary>
        /// <param name="range">A formázandó tartomány.</param>
        /// <param name="formatType">A választott számformátum típusa az enumból.</param>
        public static void SetNumberFormat(this IXLRange range, NumberFormatType formatType)
        {
            if (range is null)
                throw new ArgumentNullException(nameof(range));

            string formatString = GetFormatString(formatType);
            range.Style.NumberFormat.Format = formatString;
        }

        /// <summary>
        /// Beállítja egy adott tartomány számformátumát egy egyedi string alapján.
        /// </summary>
        /// <param name="range">A formázandó tartomány.</param>
        /// <param name="format">Az Excel számformátumát leíró string (pl. "# ##0,00", "yyyy-mm-dd", "@").</param>
        public static void SetNumberFormat(this IXLRange range, string format)
        {
            if (range is null)
                throw new ArgumentNullException(nameof(range));

            if (string.IsNullOrWhiteSpace(format))
                return;

            range.Style.NumberFormat.Format = format;
        }

        /// <summary>
        /// Segédmetódus, ami az enum értéket a megfelelő Excel formátum stringgé alakítja.
        /// </summary>
        private static string GetFormatString(NumberFormatType formatType)
        {
            switch (formatType)
            {
                case NumberFormatType.General:
                    return "General";
                case NumberFormatType.Text:
                    return "@";
                case NumberFormatType.Number:
                    return "0";
                case NumberFormatType.NumberWithTwoDecimals:
                    return "#,##0.00";
                case NumberFormatType.Currency:
                    return "#,##0.00 [$Ft-40E]"; // Magyar Forint formátum
                case NumberFormatType.Date:
                    return "yyyy.mm.dd";
                case NumberFormatType.DateTime:
                    return "yyyy.mm.dd hh:mm:ss";
                case NumberFormatType.Percentage:
                    return "0.00%";
                default:
                    return "General";
            }
        }
    }
}