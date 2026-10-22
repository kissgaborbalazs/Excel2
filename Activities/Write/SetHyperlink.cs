using System.Activities;
using System.ComponentModel;
using MPRT.Excel;

namespace MPRT.Excel.Activities.Write
{
    /// <summary>
    /// Külső URL hyperlinket állít be egy cellára, a látható szöveg megtartásával.
    /// </summary>
    [DisplayName("Set Hyperlink")]
    public class SetHyperlink : CodeActivity
    {
        [Category("Input")]
        [DisplayName("1. File Path")]
        [Description("Az .xlsx fájl teljes elérési útja.")]
        [RequiredArgument]
        public InArgument<string> FilePath { get; set; }

        [Category("Input")]
        [DisplayName("2. Sheet Name")]
        [Description("A munkalap neve (ha nem létezik, automatikusan létrejön).")]
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; }

        [Category("Input")]
        [DisplayName("3. Cell Address")]
        [Description("A hyperlinkkel ellátandó cella címe (pl. \"B7\"). A cella meglévő értéke marad a látható szöveg.")]
        [RequiredArgument]
        public InArgument<string> CellAddress { get; set; }

        [Category("Input")]
        [DisplayName("4. URL")]
        [Description("A külső URL – abszolút, sémával (pl. \"https://...\"). Csak külső cím támogatott.")]
        [RequiredArgument]
        public InArgument<string> Url { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            Excel.SetHyperlink(
                FilePath.Get(context),
                SheetName.Get(context),
                CellAddress.Get(context),
                Url.Get(context));
        }
    }
}