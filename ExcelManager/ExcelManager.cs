using System.Xml.Linq;
using ClosedXML.Excel;
using System;
using System.IO;

namespace MPRT.Excel
{
    /// <summary>
    /// Egy Excel munkafüzetet kezelő osztály, amely biztosítja a robusztus fájlkezelést és
    /// lehetővé teszi a láncolt (fluent) műveleteket.
    /// Implementálja az IDisposable interfészt a biztonságos erőforrás-kezelés érdekében.
    /// Használata 'using' blokkban javasolt.
    /// </summary>
    public partial class ExcelManager : IDisposable
    {
        private readonly XLWorkbook _workbook;
        private readonly string _filePath;

        /// <summary>
        /// Létrehoz egy új ExcelManager példányt, megnyitva vagy létrehozva a megadott fájlt.
        /// </summary>
        /// <param name="filePath">Az Excel fájl teljes elérési útvonala.</param>
        /// <exception cref="ArgumentNullException">Ha a filePath null vagy üres.</exception>
        /// <exception cref="DirectoryNotFoundException">Ha a megadott útvonal mappája nem létezik.</exception>
        /// <exception cref="ArgumentException">Ha a megadott útvonal egy létező mappára mutat.</exception>
        public ExcelManager(string filePath)
        {
            // --- Cél: Robusztus inicializálás ---
            // Mielőtt bármit csinálnánk, alaposan validáljuk a bemeneti paramétert.
            // Ezzel megelőzzük a későbbi, nehezen értelmezhető hibákat. A fejlesztő
            // azonnali, egyértelmű visszajelzést kap, ha az elérési út hibás.
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentNullException(nameof(filePath), "A fájl elérési útvonala nem lehet üres.");

            // Ellenőrizzük, hogy a célmappa létezik-e.
            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                throw new DirectoryNotFoundException($"A megadott útvonal nem létezik: {directory}");

            // Ellenőrizzük, hogy a felhasználó nem egy mappát adott-e meg fájlként.
            if (Directory.Exists(filePath))
                throw new ArgumentException("A megadott útvonal egy mappa, nem lehet fájlként megnyitni.", nameof(filePath));

            _filePath = filePath;

            // Fájl megnyitása vagy létrehozása
            if (File.Exists(_filePath))
            {
                _workbook = new XLWorkbook(_filePath);
            }
            else
            {
                _workbook = new XLWorkbook();
            }
        }

        /// <summary>
        /// Visszaad egy munkalapot a megadott néven.
        /// Ha a munkalap nem létezik, automatikusan létrehozza.
        /// </summary>
        /// <param name="sheetName">A munkalap neve.</param>
        /// <returns>A kért IXLWorksheet objektum, amin további műveletek végezhetők.</returns>
        /// <exception cref="ArgumentNullException">Ha a sheetName null vagy üres.</exception>
        public IXLWorksheet Worksheet(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
                throw new ArgumentNullException(nameof(sheetName), "A munkalap neve nem lehet üres.");

            // --- Cél: Zökkenőmentes munkalap-kezelés ---
            // Ez a metódus a "fluent" API lelke. Ahelyett, hogy kivételt dobna, ha a
            // lap nem létezik, csendben létrehozza azt. Ez lehetővé teszi a fejlesztőnek,
            // hogy megszakítás nélkül láncolja a műveleteket, pl.:
            // excel.GetWorksheet("Új Lap").Cell("A1").Value = "Működik!";
            if (!_workbook.Worksheets.TryGetWorksheet(sheetName, out var worksheet))
            {
                worksheet = _workbook.Worksheets.Add(sheetName);
            }
            return worksheet;
        }

        /// <summary>
        /// Elmenti a módosításokat és lezárja a munkafüzetet, felszabadítva az erőforrásokat.
        /// A 'using' blokk automatikusan meghívja ezt a metódust.
        /// </summary>
        public void Dispose()
        {
            // --- Cél: Garantált erőforrás-felszabadítás ---
            // Az IDisposable minta biztosítja, hogy a fájl mindig mentésre és bezárásra kerül,
            // még akkor is, ha a kódban hiba történik. A 'Dispose' metódus a 'using' blokk
            // végén fut le, garantálva a tiszta állapotot és megelőzve a fájl zárolását.
            try
            {
                // Az 'EvaluateFormulasBeforeSaving = true' beállítás arra kényszeríti a ClosedXML-t,
                // hogy a fájl lemezre írása előtt futtassa le és számolja ki az összes képletet.
                // Ez biztosítja, hogy a mentett fájl mindig a friss, aktuális értékeket tartalmazza
                ClosedXML.Excel.SaveOptions saveOptions = new ClosedXML.Excel.SaveOptions { EvaluateFormulasBeforeSaving = true };
                _workbook.SaveAs(_filePath, saveOptions);
            }
            finally
            {
                _workbook.Dispose();
            }
        }
    }
}