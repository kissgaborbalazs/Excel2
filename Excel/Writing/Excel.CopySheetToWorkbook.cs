using System.IO;
using ClosedXML.Excel;
using System;

namespace MPRT.Excel
{
    /// <summary>
    /// Egyszerű, statikus segédmetódusokat biztosít a leggyakoribb Excel műveletekhez.
    /// (Ez a class partial, a metódusok külön fájlokban vannak a jobb szervezettség érdekében.)
    /// </summary>
    public static partial class Excel
    {
        /// <summary>
        /// Átmásol egy munkalapot egyik Excel fájlból egy másikba.
        /// </summary>
        /// <param name="sourceFilePath">A forrás Excel fájl teljes elérési útvonala.</param>
        /// <param name="sourceSheetName">A másolandó munkalap neve a forrás fájlban.</param>
        /// <param name="destinationFilePath">A cél Excel fájl teljes elérési útvonala. Ha nem létezik, létrehozza.</param>
        /// <param name="destinationSheetName">A munkalap neve a cél fájlban.</param>
        public static void CopySheetToWorkbook(string sourceFilePath, string sourceSheetName, string destinationFilePath, string destinationSheetName)
        {
            // Cél: Két külön fájl kezelése.
            // Mivel ez a művelet két munkafüzetet érint egyszerre, eltérünk a szabványos ExcelManager
            // használatától, ami egyetlen fájl életciklusára van tervezve. Itt közvetlenül
            // kezeljük a ClosedXML munkafüzet objektumait a maximális kontroll érdekében.

            XLWorkbook sourceWorkbook = null;
            XLWorkbook destinationWorkbook = null;

            try
            {
                // 1. Forrásfájl megnyitása
                if (!File.Exists(sourceFilePath))
                    throw new FileNotFoundException("A forrásfájl nem található.", sourceFilePath);
                
                sourceWorkbook = new XLWorkbook(sourceFilePath);
                
                // 2. Forrás munkalap lekérése
                if (!sourceWorkbook.Worksheets.TryGetWorksheet(sourceSheetName, out var sourceSheet))
                    throw new ArgumentException($"A '{sourceSheetName}' nevű munkalap nem létezik a forrásfájlban.", nameof(sourceSheetName));

                // 3. Célfájl megnyitása vagy létrehozása
                destinationWorkbook = File.Exists(destinationFilePath) 
                    ? new XLWorkbook(destinationFilePath) 
                    : new XLWorkbook();

                if (string.IsNullOrWhiteSpace(destinationSheetName))
                    destinationSheetName = sourceSheetName;
                // 4. A másolás végrehajtása
                // A ClosedXML beépített CopyTo metódusa elvégzi a munka nehezét:
                // átmásolja a cellákat, stílusokat, formátumokat stb.
                sourceSheet.CopyTo(destinationWorkbook, destinationSheetName);

                // 5. Célfájl mentése
                destinationWorkbook.SaveAs(destinationFilePath);
            }
            finally
            {
                // Cél: Garantált erőforrás-felszabadítás.
                // A finally blokk biztosítja, hogy mindkét munkafüzet erőforrásai
                // felszabaduljanak, még akkor is, ha a művelet során hiba történik.
                sourceWorkbook?.Dispose();
                destinationWorkbook?.Dispose();
            }
        }
    }
}