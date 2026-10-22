using ClosedXML.Excel;

namespace MPRT.Excel
{
    /// <summary>
    /// Egy stílust leíró osztály, amely a formázási beállításokat tartalmazza.
    /// A nullable típusok (?) lehetővé teszik, hogy csak azokat a tulajdonságokat
    /// kelljen beállítani, amiket ténylegesen módosítani szeretnénk.
    /// </summary>
    public class Style
    {
        // Betűstílusok
        public bool? IsBold { get; set; }
        public bool? IsItalic { get; set; }
        public double? FontSize { get; set; }

        // Színek
        public XLColor? BackgroundColor { get; set; }
        public XLColor? FontColor { get; set; }
        
        // Szegélyek
        public XLBorderStyleValues? OutsideBorder { get; set; }
        public XLBorderStyleValues? InsideBorder { get; set; }
        public XLColor? BorderColor { get; set; }
    }
}
