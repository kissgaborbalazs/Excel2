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
        /// Beállítja egy adott tartomány stílusát a megadott Style objektum alapján.
        /// </summary>
        /// <param name="range">A formázandó tartomány.</param>
        /// <param name="style">A stílusbeállításokat tartalmazó objektum.</param>
        public static void SetStyle(this IXLRange range, Style style)
        {
            // Cél: Rugalmas stílusbeállítás.
            // A metódus végigmegy a style objektum tulajdonságain, és csak azokat
            // alkalmazza, amelyeknek a felhasználó értéket adott (nem null).
            // Ez lehetővé teszi a stílusok részleges módosítását.
            
            if (style == null)
                return;

            //IXLStyle xlStyle = range.Style;

            // Betű formázása
            if (style.IsBold.HasValue)
                range.Style.Font.Bold = style.IsBold.Value;

            if (style.IsItalic.HasValue)
                range.Style.Font.Italic = style.IsItalic.Value;

            if (style.FontSize.HasValue)
                range.Style.Font.FontSize = style.FontSize.Value;

            // Színek beállítása
            if (!(style.BackgroundColor is null) && style.BackgroundColor.HasValue)
                range.Style.Fill.BackgroundColor = (XLColor)style.BackgroundColor;

            if (!(style.FontColor is null) && style.FontColor.HasValue)
                range.Style.Font.FontColor = (XLColor)style.FontColor;

            // Szegélyek formázása
            var border = range.Style.Border;

            if (style.OutsideBorder.HasValue)
            {
                //range.Style.Border.SetOutsideBorder((XLBorderStyleValues)style.OutsideBorder.Value);
                border.TopBorder = style.OutsideBorder.Value;
                border.BottomBorder = style.OutsideBorder.Value;
                border.LeftBorder = style.OutsideBorder.Value;
                border.RightBorder = style.OutsideBorder.Value;
            }

            if (style.InsideBorder.HasValue)
            {
                border.SetInsideBorder(style.InsideBorder.Value);
            }
            
            if (!(style.BorderColor is null) && style.BorderColor.HasValue)
            {
                // A színt alkalmazzuk mind a külső, mind a belső szegélyekre,
                // ha azok definiálva vannak.
                if (style.OutsideBorder.HasValue)
                {
                    //range.Style.Border.SetOutsideBorderColor((XLColor)style.BorderColor);
                    border.TopBorderColor = style.BorderColor;
                    border.BottomBorderColor = style.BorderColor;
                    border.LeftBorderColor = style.BorderColor;
                    border.RightBorderColor = style.BorderColor;
                }
                if (style.InsideBorder.HasValue)
                    range.Style.Border.SetInsideBorderColor((XLColor)style.BorderColor);
            }
        }
    }
}