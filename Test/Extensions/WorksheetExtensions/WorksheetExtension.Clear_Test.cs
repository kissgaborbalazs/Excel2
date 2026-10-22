using System;
using System.IO;
using UiPath.CodedWorkflows;
using ClosedXML.Excel;
using System.Diagnostics;
using System.Linq;
using System.Collections.Generic;

// A tesztelendő osztály névtere.
using MPRT.Excel;

namespace MPRT.Excel.Tests
{
    public class WorksheetExtensions_Clear_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorksheetExtClearTests", Guid.NewGuid().ToString());
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

            runTest("Clear(false): Csak a tartalmat törli, a stílust meghagyja", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    var cell = ws.Cell("A1");
                    cell.Value = "Teszt Adat";
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.Yellow;

                    // Act
                    WorksheetExtensions.Clear(ws, false); // clearStyles = false (alapértelmezett)

                    // Assert
                    if (!cell.IsEmpty())
                        throw new Exception($"A cella tartalma nem lett törölve. Várt: üres, Kapott: '{cell.Value}'");
                    if (!cell.Style.Font.Bold)
                        throw new Exception($"A cella 'Bold' stílusa hibásan törlődött. Várt: true, Kapott: false");
                    if (cell.Style.Fill.BackgroundColor.Color != XLColor.Yellow.Color)
                        throw new Exception($"A cella háttérszíne hibásan törlődött. Várt: {XLColor.Yellow.Color}, Kapott: {cell.Style.Fill.BackgroundColor.Color}");
                }
            });

            runTest("Clear(true): A tartalmat és a stílust is törli", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    var cell = ws.Cell("B2");
                    cell.Value = "Minden törlődik";
                    cell.Style.Font.Italic = true;
                    cell.Style.Fill.BackgroundColor = XLColor.LightBlue;

                    // Act
                    WorksheetExtensions.Clear(ws, true);

                    // Assert
                    if (!cell.IsEmpty())
                        throw new Exception($"A cella tartalma nem lett törölve. Várt: üres, Kapott: '{cell.Value}'");
                    if (cell.Style.Font.Italic)
                        throw new Exception($"A cella 'Italic' stílusa nem lett törölve. Várt: false, Kapott: true");
                    if (cell.Style.Fill.BackgroundColor.Color != XLColor.Transparent.Color)
                        throw new Exception($"A cella háttérszíne nem lett törölve. Várt: {XLColor.Transparent.Color}, Kapott: {cell.Style.Fill.BackgroundColor.Color}");
                }
            });

            runTest("Clear: Üres munkalapon nem okoz hibát", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");

                    // Act & Assert (nem dobhat kivételt)
                    WorksheetExtensions.Clear(ws);
                    WorksheetExtensions.Clear(ws, true);
                }
            });

            runTest("Clear(false): Több cella és képlet tartalmát is törli", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Szöveg";
                    ws.Cell("B2").Value = 123;
                    ws.Cell("C3").FormulaA1 = "=A1"; // Képlet

                    // Act
                    WorksheetExtensions.Clear(ws, false);

                    // Assert
                    if (ws.LastCellUsed() != null)
                        throw new Exception($"A munkalap nem lett teljesen kiürítve. Utolsó cella: {ws.LastCellUsed().Address}");
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