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
    public class WorksheetExtensions_FreezePanes_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorksheetExtFreezePanesTests", Guid.NewGuid().ToString());
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

            runTest("FreezePanes: Helyesen rögzíti a sorokat és oszlopokat is", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");

                    // Act
                    WorksheetExtensions.FreezePanes(ws, 2, 3);

                    // Assert
                    if (ws.SheetView.SplitRow != 2)
                        throw new Exception($"A rögzített sorok száma hibás. Várt: 2, Kapott: {ws.SheetView.SplitRow}");
                    if (ws.SheetView.SplitColumn != 3)
                        throw new Exception($"A rögzített oszlopok száma hibás. Várt: 3, Kapott: {ws.SheetView.SplitColumn}");
                }
            });

            runTest("FreezePanes: Csak sorokat rögzít, ha az oszlop 0", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");

                    // Act
                    WorksheetExtensions.FreezePanes(ws, 1, 0);

                    // Assert
                    if (ws.SheetView.SplitRow != 1)
                        throw new Exception($"A rögzített sorok száma hibás. Várt: 1, Kapott: {ws.SheetView.SplitRow}");
                    if (ws.SheetView.SplitColumn != 0)
                        throw new Exception($"A rögzített oszlopok száma hibás. Várt: 0, Kapott: {ws.SheetView.SplitColumn}");
                }
            });

            runTest("FreezePanes: Csak oszlopokat rögzít, ha a sor 0", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");

                    // Act
                    WorksheetExtensions.FreezePanes(ws, 0, 4);

                    // Assert
                    if (ws.SheetView.SplitRow != 0)
                        throw new Exception($"A rögzített sorok száma hibás. Várt: 0, Kapott: {ws.SheetView.SplitRow}");
                    if (ws.SheetView.SplitColumn != 4)
                        throw new Exception($"A rögzített oszlopok száma hibás. Várt: 4, Kapott: {ws.SheetView.SplitColumn}");
                }
            });

            runTest("FreezePanes: (0, 0) értékekkel feloldja a rögzítést", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.SheetView.Freeze(1, 1); // Előzetes rögzítés

                    // Act
                    WorksheetExtensions.FreezePanes(ws, 0, 0);

                    // Assert
                    if (ws.SheetView.SplitRow != 0)
                        throw new Exception($"A sorrögzítés nem lett feloldva. Várt: 0, Kapott: {ws.SheetView.SplitRow}");
                    if (ws.SheetView.SplitColumn != 0)
                        throw new Exception($"Az oszloprögzítés nem lett feloldva. Várt: 0, Kapott: {ws.SheetView.SplitColumn}");
                }
            });

            runTest("FreezePanes: Negatív értékeket figyelmen kívül hagy, nem okoz hibát", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.SheetView.Freeze(1, 1); // Előzetes rögzítés

                    // Act & Assert (nem dobhat kivételt)
                    WorksheetExtensions.FreezePanes(ws, -1, -5);
                    
                    // A rögzítésnek változatlannak kell maradnia
                    if (ws.SheetView.SplitRow != 1)
                        throw new Exception($"A sorrögzítés hibásan megváltozott. Várt: 1, Kapott: {ws.SheetView.SplitRow}");
                    if (ws.SheetView.SplitColumn != 1)
                        throw new Exception($"Az oszloprögzítés hibásan megváltozott. Várt: 1, Kapott: {ws.SheetView.SplitColumn}");
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
