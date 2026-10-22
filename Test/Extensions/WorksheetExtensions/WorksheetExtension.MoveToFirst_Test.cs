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
    public class WorksheetExtensions_MoveSheetToFirst_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorksheetExtMoveToFirstTests", Guid.NewGuid().ToString());
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

            runTest("MoveSheetToFirst: A harmadik lapot az első helyre mozgatja", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    workbook.Worksheets.Add("Lap 1");
                    workbook.Worksheets.Add("Lap 2");
                    var sheetToMove = workbook.Worksheets.Add("Lap 3");

                    // Act
                    WorksheetExtensions.MoveSheetToFirst(sheetToMove);

                    // Assert
                    // JAVÍTVA: Direktebb ellenőrzés a pozíció alapján.
                    var firstSheetName = workbook.Worksheet(1).Name;
                    if (firstSheetName != "Lap 3")
                    {
                        throw new Exception($"Az első pozícióban lévő lap neve hibás. Várt: 'Lap 3', Kapott: '{firstSheetName}'");
                    }
                }
            });

            runTest("MoveSheetToFirst: Ha a lap már az első helyen van, nem történik változás", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var sheetToMove = workbook.Worksheets.Add("Lap 1");
                    workbook.Worksheets.Add("Lap 2");
                    
                    var initialOrder = workbook.Worksheets.Select(ws => ws.Name).ToList();

                    // Act
                    WorksheetExtensions.MoveSheetToFirst(sheetToMove);

                    // Assert
                    var finalOrder = workbook.Worksheets.Select(ws => ws.Name).ToList();
                    if (!initialOrder.SequenceEqual(finalOrder))
                    {
                        throw new Exception($"A munkalapok sorrendje hibásan megváltozott. Várt: {string.Join(", ", initialOrder)}, Kapott: {string.Join(", ", finalOrder)}");
                    }
                }
            });

            runTest("MoveSheetToFirst: Egyetlen munkalap esetén nem okoz hibát", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var sheetToMove = workbook.Worksheets.Add("Egyetlen Lap");

                    // Act & Assert (nem dobhat kivételt)
                    WorksheetExtensions.MoveSheetToFirst(sheetToMove);

                    if (workbook.Worksheet(1).Name != "Egyetlen Lap")
                    {
                        throw new Exception("Az egyetlen munkalap neve vagy pozíciója megváltozott.");
                    }
                }
            });

            runTest("MoveSheetToFirst: Null worksheet esetén ArgumentNullException-t dob", () =>
            {
                // Arrange
                IXLWorksheet nullSheet = null;

                // Act & Assert
                try
                {
                    WorksheetExtensions.MoveSheetToFirst(nullSheet);
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
