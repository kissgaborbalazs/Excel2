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
    public class WorksheetExtensions_UnhideColumn_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorksheetExtUnhideColTests", Guid.NewGuid().ToString());
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

            runTest("UnhideColumn: Sikeresen felfed egy elrejtett oszlopot", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    var columnToUnhide = ws.Column(3); // C oszlop
                    columnToUnhide.Hide();

                    // Act
                    WorksheetExtensions.UnhideColumn(ws, 3);

                    // Assert
                    if (columnToUnhide.IsHidden)
                        throw new Exception($"Az oszlop nem lett felfedve. Várt: IsHidden=false, Kapott: {columnToUnhide.IsHidden}");
                }
            });

            runTest("UnhideColumn: Már látható oszlopon nem változtat", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    var column = ws.Column(1); // A oszlop

                    // Act
                    WorksheetExtensions.UnhideColumn(ws, 1);

                    // Assert
                    if (column.IsHidden)
                        throw new Exception($"Egy látható oszlop hibásan rejtve lett. Várt: IsHidden=false, Kapott: {column.IsHidden}");
                }
            });

            runTest("UnhideColumn: Érvénytelen (nulla vagy negatív) oszlopindexet figyelmen kívül hagy", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Column(1).Hide();

                    // Act & Assert (nem dobhat kivételt)
                    WorksheetExtensions.UnhideColumn(ws, 0);
                    WorksheetExtensions.UnhideColumn(ws, -5);

                    if (!ws.Column(1).IsHidden)
                        throw new Exception("Az 'A' oszlop hibásan fel lett fedve érvénytelen indexek hatására.");
                }
            });

            runTest("UnhideColumn: Null worksheet esetén ArgumentNullException-t dob", () =>
            {
                // Arrange
                IXLWorksheet nullSheet = null;

                // Act & Assert
                try
                {
                    WorksheetExtensions.UnhideColumn(nullSheet, 1);
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
