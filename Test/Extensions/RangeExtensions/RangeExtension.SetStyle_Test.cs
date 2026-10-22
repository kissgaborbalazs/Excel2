using System;
using System.IO;
using UiPath.CodedWorkflows;
using ClosedXML.Excel;
using System.Diagnostics;
using System.Linq;

// A tesztelendő osztály névtere.
using MPRT.Excel;

namespace MPRT.Excel.Tests
{
    public class RangeExtensions_SetStyle_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "RangeExtensionsSetStyleTests", Guid.NewGuid().ToString());
            Directory.CreateDirectory(testDir);
            Log($"Temporary test directory created at: {testDir}");

            Action<string, Action> runTest = (testName, testAction) =>
            {
                total++;
                Log($"\n▶️ RUNNING: {testName}");
                try
                {
                    testAction();
                    Log($"✅ PASS: {testName}");
                    passed++;
                }
                catch (Exception ex)
                {
                    Log($"❌ FAIL: {testName}\n   Hiba: {ex.Message}\n   Részletek: {ex.StackTrace}");
                    failed++;
                }
            };
            
            // =================================================================
            // TESZTESETEK DEFINÍCIÓJA
            // =================================================================

            runTest("SetStyle: Minden stílust helyesen alkalmaz", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("TestSheet");
                    var range = worksheet.Range("A1:C3"); // MÓDOSÍTVA
                    var fullStyle = new Style
                    {
                        IsBold = true,
                        IsItalic = true,
                        FontSize = 14,
                        BackgroundColor = XLColor.LightGray,
                        FontColor = XLColor.Red,
                        OutsideBorder = XLBorderStyleValues.Thick,
                        InsideBorder = XLBorderStyleValues.Dotted,
                        BorderColor = XLColor.Blue
                    };

                    // Act
                    RangeExtensions.SetStyle(range, fullStyle);

                    // Assert - Általános stílusok
                    if (!range.Style.Font.Bold) throw new Exception($"A 'Bold' stílus hibás. Várt: true, Kapott: {range.Style.Font.Bold}");
                    if (!range.Style.Font.Italic) throw new Exception($"Az 'Italic' stílus hibás. Várt: true, Kapott: {range.Style.Font.Italic}");
                    if (range.Style.Font.FontSize != 14) throw new Exception($"A betűméret hibás. Várt: 14, Kapott: {range.Style.Font.FontSize}");
                    if (range.Style.Fill.BackgroundColor.Color != XLColor.LightGray.Color) throw new Exception($"A háttérszín hibás. Várt: {XLColor.LightGray.Color}, Kapott: {range.Style.Fill.BackgroundColor.Color}");
                    if (range.Style.Font.FontColor.Color != XLColor.Red.Color) throw new Exception($"A betűszín hibás. Várt: {XLColor.Red.Color}, Kapott: {range.Style.Font.FontColor.Color}");

                    // Assert - Részletes szegélyellenőrzés
                    // Helper a szegélyek ellenőrzéséhez
                    Action<IXLCell, XLBorderStyleValues, XLBorderStyleValues, XLBorderStyleValues, XLBorderStyleValues> checkBorders = 
                        (cell, top, bottom, left, right) =>
                    {
                        if (cell.Style.Border.TopBorder != top) throw new Exception($"Hiba a(z) {cell.Address} cellánál. Felső szegély. Várt: {top}, Kapott: {cell.Style.Border.TopBorder}");
                        if (cell.Style.Border.BottomBorder != bottom) throw new Exception($"Hiba a(z) {cell.Address} cellánál. Alsó szegély. Várt: {bottom}, Kapott: {cell.Style.Border.BottomBorder}");
                        if (cell.Style.Border.LeftBorder != left) throw new Exception($"Hiba a(z) {cell.Address} cellánál. Bal szegély. Várt: {left}, Kapott: {cell.Style.Border.LeftBorder}");
                        if (cell.Style.Border.RightBorder != right) throw new Exception($"Hiba a(z) {cell.Address} cellánál. Jobb szegély. Várt: {right}, Kapott: {cell.Style.Border.RightBorder}");
                        if (cell.Style.Border.TopBorderColor.Color != XLColor.Blue.Color) throw new Exception($"Hiba a(z) {cell.Address} cellánál. A szegélyszín hibás.");
                    };

                    // Felső sor
                    checkBorders(worksheet.Cell("A1"), XLBorderStyleValues.Thick, XLBorderStyleValues.Dotted, XLBorderStyleValues.Thick, XLBorderStyleValues.Dotted);
                    checkBorders(worksheet.Cell("B1"), XLBorderStyleValues.Thick, XLBorderStyleValues.Dotted, XLBorderStyleValues.Dotted, XLBorderStyleValues.Dotted);
                    checkBorders(worksheet.Cell("C1"), XLBorderStyleValues.Thick, XLBorderStyleValues.Dotted, XLBorderStyleValues.Dotted, XLBorderStyleValues.Thick);
                    
                    // Középső sor
                    checkBorders(worksheet.Cell("A2"), XLBorderStyleValues.Dotted, XLBorderStyleValues.Dotted, XLBorderStyleValues.Thick, XLBorderStyleValues.Dotted);
                    checkBorders(worksheet.Cell("B2"), XLBorderStyleValues.Dotted, XLBorderStyleValues.Dotted, XLBorderStyleValues.Dotted, XLBorderStyleValues.Dotted);
                    checkBorders(worksheet.Cell("C2"), XLBorderStyleValues.Dotted, XLBorderStyleValues.Dotted, XLBorderStyleValues.Dotted, XLBorderStyleValues.Thick);

                    // Alsó sor
                    checkBorders(worksheet.Cell("A3"), XLBorderStyleValues.Dotted, XLBorderStyleValues.Thick, XLBorderStyleValues.Thick, XLBorderStyleValues.Dotted);
                    checkBorders(worksheet.Cell("B3"), XLBorderStyleValues.Dotted, XLBorderStyleValues.Thick, XLBorderStyleValues.Dotted, XLBorderStyleValues.Dotted);
                    checkBorders(worksheet.Cell("C3"), XLBorderStyleValues.Dotted, XLBorderStyleValues.Thick, XLBorderStyleValues.Dotted, XLBorderStyleValues.Thick);
                }
            });

            runTest("SetStyle: Csak a megadott stílusokat alkalmazza (részleges stílus)", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("TestSheet");
                    var range = worksheet.Range("C3:D4");
                    // Kezdeti stílus beállítása
                    range.Style.Font.Bold = true;
                    range.Style.Fill.BackgroundColor = XLColor.Yellow;

                    var partialStyle = new Style
                    {
                        IsItalic = true, // Ezt hozzáadja
                        BackgroundColor = XLColor.Green // Ezt felülírja
                        // A Bold-ot nem bántja
                    };

                    // Act
                    RangeExtensions.SetStyle(range, partialStyle);

                    // Assert
                    if (!range.Style.Font.Bold) throw new Exception($"Az eredeti 'Bold' stílus elveszett. Várt: true, Kapott: {range.Style.Font.Bold}");
                    if (!range.Style.Font.Italic) throw new Exception($"Az 'Italic' stílus nem lett alkalmazva. Várt: true, Kapott: {range.Style.Font.Italic}");
                    if (range.Style.Fill.BackgroundColor.Color != XLColor.Green.Color) throw new Exception($"A háttérszín nem lett felülírva. Várt: {XLColor.Green.Color}, Kapott: {range.Style.Fill.BackgroundColor.Color}");
                }
            });

            runTest("SetStyle: Üres stílus objektummal nem változtat semmin", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("TestSheet");
                    var range = worksheet.Range("E5:F6");
                    range.Style.Font.Bold = true;
                    range.Style.Fill.BackgroundColor = XLColor.Red;

                    var emptyStyle = new Style(); // Minden tulajdonsága null

                    // Act
                    RangeExtensions.SetStyle(range, emptyStyle);

                    // Assert
                    if (!range.Style.Font.Bold) throw new Exception($"A 'Bold' stílus hibásan megváltozott. Várt: true, Kapott: {range.Style.Font.Bold}");
                    if (range.Style.Fill.BackgroundColor.Color != XLColor.Red.Color) throw new Exception($"A háttérszín hibásan megváltozott. Várt: {XLColor.Red.Color}, Kapott: {range.Style.Fill.BackgroundColor.Color}");
                }
            });

            runTest("SetStyle: Stílusok kikapcsolása (false értékek)", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("TestSheet");
                    var range = worksheet.Range("G7:H8");
                    range.Style.Font.Bold = true;
                    range.Style.Font.Italic = true;

                    var turnOffStyle = new Style
                    {
                        IsBold = false,
                        IsItalic = false
                    };

                    // Act
                    RangeExtensions.SetStyle(range, turnOffStyle);

                    // Assert
                    if (range.Style.Font.Bold) throw new Exception($"A 'Bold' stílus nem lett kikapcsolva. Várt: false, Kapott: {range.Style.Font.Bold}");
                    if (range.Style.Font.Italic) throw new Exception($"Az 'Italic' stílus nem lett kikapcsolva. Várt: false, Kapott: {range.Style.Font.Italic}");
                }
            });

            runTest("SetStyle: Szegélyszín csak akkor érvényesül, ha a szegély is be van állítva", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("TestSheet");
                    var rangeWithBorder = worksheet.Range("A1:B2");
                    var rangeWithoutBorder = worksheet.Range("D4:E5");

                    var styleWithColor = new Style
                    {
                        OutsideBorder = XLBorderStyleValues.Thin,
                        BorderColor = XLColor.Magenta
                    };
                    
                    var styleOnlyColor = new Style
                    {
                        BorderColor = XLColor.Cyan
                    };

                    // Act
                    RangeExtensions.SetStyle(rangeWithBorder, styleWithColor);
                    RangeExtensions.SetStyle(rangeWithoutBorder, styleOnlyColor);

                    // Assert
                    if (rangeWithBorder.Style.Border.TopBorderColor.Color != XLColor.Magenta.Color) 
                        throw new Exception($"A szegélyszín hibás. Várt: {XLColor.Magenta.Color}, Kapott: {rangeWithBorder.Style.Border.TopBorderColor.Color}");
                    
                    if (rangeWithoutBorder.Style.Border.TopBorderColor.Color != XLColor.Black.Color)
                        throw new Exception($"A szegélyszín hibásan lett beállítva szegély nélkül. Várt: {XLColor.Black.Color}, Kapott: {rangeWithoutBorder.Style.Border.TopBorderColor.Color}");
                    if (rangeWithoutBorder.Style.Border.TopBorder != XLBorderStyleValues.None)
                        throw new Exception($"A szegély hibásan lett beállítva, amikor csak színt adtunk meg. Várt: {XLBorderStyleValues.None}, Kapott: {rangeWithoutBorder.Style.Border.TopBorder}");
                }
            });

            // =================================================================
            // TESZTEK ÖSSZEGZÉSE
            // =================================================================
            
            Log("\n---");
            Log($"🟢 Sikeres: {passed}, 🔴 Sikertelen: {failed} / {total} teszt");
            
            // Cleanup the temporary directory
            try
            {
                Directory.Delete(testDir, true);
                Log($"Temporary test directory successfully deleted.");
            }
            catch (Exception ex)
            {
                Log($"Warning: Could not delete temporary directory {testDir}. Hiba: {ex.Message}");
            }

            // Végső ellenőrzés a UiPath Test Suite számára
            testing.VerifyExpression(failed == 0, $"Tesztek összegzése: {passed} sikeres, {failed} sikertelen.", true, null, false, false);
        }
    }
}