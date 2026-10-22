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
    public class WorksheetExtensions_HideRow_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorksheetExtHideRowTests", Guid.NewGuid().ToString());
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

            runTest("HideRow: Sikeresen elrejt egy sort", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A5").Value = "Ezt el kell rejteni"; // 5. sor

                    // Act
                    WorksheetExtensions.HideRow(ws, 5);

                    // Assert
                    if (!ws.Row(5).IsHidden)
                        throw new Exception($"A sor nem lett elrejtve. Várt: IsHidden=true, Kapott: {ws.Row(5).IsHidden}");
                }
            });

            runTest("HideRow: Nem hatással más sorokra", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Látható";
                    ws.Cell("A2").Value = "Rejtett";
                    ws.Cell("A3").Value = "Látható";

                    // Act
                    WorksheetExtensions.HideRow(ws, 2);

                    // Assert
                    if (ws.Row(1).IsHidden)
                        throw new Exception($"Az 1. sor hibásan rejtve lett. Várt: IsHidden=false, Kapott: {ws.Row(1).IsHidden}");
                    if (ws.Row(3).IsHidden)
                        throw new Exception($"A 3. sor hibásan rejtve lett. Várt: IsHidden=false, Kapott: {ws.Row(3).IsHidden}");
                }
            });

            runTest("HideRow: Érvénytelen (nulla vagy negatív) sorindexet figyelmen kívül hagy", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Teszt";

                    // Act & Assert (nem dobhat kivételt)
                    WorksheetExtensions.HideRow(ws, 0);
                    WorksheetExtensions.HideRow(ws, -5);

                    if (ws.Row(1).IsHidden)
                        throw new Exception("Az 1. sor hibásan rejtve lett érvénytelen indexek hatására.");
                }
            });

            runTest("HideRow: Üres munkalapon nem okoz hibát", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");

                    // Act & Assert (nem dobhat kivételt)
                    WorksheetExtensions.HideRow(ws, 1);
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
