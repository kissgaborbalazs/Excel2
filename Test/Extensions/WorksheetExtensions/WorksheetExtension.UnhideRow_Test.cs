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
    public class WorksheetExtensions_UnhideRow_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorksheetExtUnhideRowTests", Guid.NewGuid().ToString());
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

            runTest("UnhideRow: Sikeresen felfed egy elrejtett sort", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    var rowToUnhide = ws.Row(5);
                    rowToUnhide.Hide();

                    // Act
                    WorksheetExtensions.UnhideRow(ws, 5);

                    // Assert
                    if (rowToUnhide.IsHidden)
                        throw new Exception($"A sor nem lett felfedve. Várt: IsHidden=false, Kapott: {rowToUnhide.IsHidden}");
                }
            });

            runTest("UnhideRow: Már látható soron nem változtat", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    var row = ws.Row(1);

                    // Act
                    WorksheetExtensions.UnhideRow(ws, 1);

                    // Assert
                    if (row.IsHidden)
                        throw new Exception($"Egy látható sor hibásan rejtve lett. Várt: IsHidden=false, Kapott: {row.IsHidden}");
                }
            });

            runTest("UnhideRow: Érvénytelen (nulla vagy negatív) sorindexet figyelmen kívül hagy", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Row(1).Hide();

                    // Act & Assert (nem dobhat kivételt)
                    WorksheetExtensions.UnhideRow(ws, 0);
                    WorksheetExtensions.UnhideRow(ws, -5);

                    if (!ws.Row(1).IsHidden)
                        throw new Exception("Az 1. sor hibásan fel lett fedve érvénytelen indexek hatására.");
                }
            });

            runTest("UnhideRow: Null worksheet esetén ArgumentNullException-t dob", () =>
            {
                // Arrange
                IXLWorksheet nullSheet = null;

                // Act & Assert
                try
                {
                    WorksheetExtensions.UnhideRow(nullSheet, 1);
                    throw new Exception("Nem dobódott a várt ArgumentNullException.");
                }
                catch (ArgumentNullException)
                {
                    /* Test Passed */
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
