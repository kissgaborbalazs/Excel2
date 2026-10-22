using System.Activities;
using System.ComponentModel;
using MPRT.Excel;
using ClosedXML.Excel;
using System;

namespace MPRT.Excel.Activities.Format
{
    /// <summary>
    /// A felhasználóbarát, legördülő menüben megjelenő alapszínek listája.
    /// </summary>
    public enum BasicColor
    {
        None,
        Black,
        White,
        Red,
        Green,
        Blue,
        Yellow,
        Gray,
        LightGray,
        Transparent
    }

    [DisplayName("Set Range Style")]
    [Description("Sets the style for a specified range using individual properties.")]
    public class SetRangeStyleSimple : CodeActivity
    {
        #region Inputs
        [Category("Input")]
        [RequiredArgument]
        [DisplayName("1. File Path")]
        [Description("The full path to the Excel file.")]
        public InArgument<string> FilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("2. Sheet Name")]
        [Description("The name of the worksheet.")]
        public InArgument<string> SheetName { get; set; }

        [Category("Input")]
        [DisplayName("3. Range")]
        [Description("The range to format (e.g., \"A1:C10\").")]
        public InArgument<string> Range { get; set; }
        #endregion

        #region Style Properties
        [Category("Style Options")]
        [DisplayName("Is Bold")]
        public bool IsBold { get; set; } = false;

        [Category("Style Options")]
        [DisplayName("Is Italic")]
        public bool IsItalic { get; set; } = false;

        [Category("Style Options")]
        [DisplayName("Font Size")]
        public InArgument<double?> FontSize { get; set; }

        [Category("Style Options")]
        [DisplayName("Background Color")]
        [DefaultValue(BasicColor.None)]
        public BasicColor BackgroundColor { get; set; } = BasicColor.None;

        [Category("Style Options")]
        [DisplayName("Font Color")]
        [DefaultValue(BasicColor.None)]
        public BasicColor FontColor { get; set; } = BasicColor.None;

        [Category("Style Options")]
        [DisplayName("Outside Border")]
        [DefaultValue(XLBorderStyleValues.None)]
        public XLBorderStyleValues OutsideBorder { get; set; } = XLBorderStyleValues.None;

        [Category("Style Options")]
        [DisplayName("Inside Border")]
        [DefaultValue(XLBorderStyleValues.None)]
        public XLBorderStyleValues InsideBorder { get; set; } = XLBorderStyleValues.None;

        [Category("Style Options")]
        [DisplayName("Border Color")]
        [DefaultValue(BasicColor.None)]
        public BasicColor BorderColor { get; set; } = BasicColor.None;
        #endregion

        protected override void Execute(CodeActivityContext context)
        {
            var filePath = FilePath.Get(context);
            var sheetName = SheetName.Get(context);
            var range = Range.Get(context);

            var style = new Style();

            //if (IsBold.Get(context).HasValue) style.IsBold = IsBold.Get(context);
            style.IsBold = IsBold;
            //if (IsItalic.Get(context).HasValue) style.IsItalic = IsItalic.Get(context);
            style.IsItalic = IsItalic;
            if (FontSize.Get(context).HasValue) style.FontSize = FontSize.Get(context);

            // JAVÍTÁS: A kiválasztott BasicColor enum értéket átalakítjuk a háttérben szükséges XLColor-rá.
            if (BackgroundColor != BasicColor.None) style.BackgroundColor = ConvertToXLColor(BackgroundColor);
            if (FontColor != BasicColor.None) style.FontColor = ConvertToXLColor(FontColor);
            
            if (OutsideBorder != XLBorderStyleValues.None) style.OutsideBorder = OutsideBorder;
            if (InsideBorder != XLBorderStyleValues.None) style.InsideBorder = InsideBorder;

            bool borderIsSet = (OutsideBorder != XLBorderStyleValues.None || InsideBorder != XLBorderStyleValues.None);

            if (borderIsSet)
            {
                style.BorderColor = (BorderColor != BasicColor.None) ? ConvertToXLColor(BorderColor) : XLColor.Black;
            }

            Excel.SetRangeStyle(filePath, sheetName, range, style);
        }

        /// <summary>
        /// Segédmetódus, ami a felhasználóbarát BasicColor enum-ot átalakítja
        /// a ClosedXML által megkövetelt XLColor objektummá.
        /// </summary>
        private XLColor ConvertToXLColor(BasicColor color)
        {
            switch (color)
            {
                case BasicColor.Black: return XLColor.Black;
                case BasicColor.White: return XLColor.White;
                case BasicColor.Red: return XLColor.Red;
                case BasicColor.Green: return XLColor.Green;
                case BasicColor.Blue: return XLColor.Blue;
                case BasicColor.Yellow: return XLColor.Yellow;
                case BasicColor.Gray: return XLColor.Gray;
                case BasicColor.LightGray: return XLColor.LightGray;
                case BasicColor.Transparent: return XLColor.Transparent;
                default: return XLColor.NoColor; // Biztonságos alapértelmezett
            }
        }
    }
}