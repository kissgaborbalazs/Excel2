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
    public class WorksheetExtensions_SetSheetPosition_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorksheetExtSetPosTests", Guid.NewGuid().ToString());
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

            runTest("SetSheetPosition: Munkalapot érvényes pozícióba mozgat", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    workbook.Worksheets.Add("Lap 1");
                    var sheetToMove = workbook.Worksheets.Add("Lap 2");
                    workbook.Worksheets.Add("Lap 3");

                    // Act
                    WorksheetExtensions.SetSheetPosition(sheetToMove, 1);

                    // Assert
                    if (workbook.Worksheet(1).Name != "Lap 2")
                        throw new Exception($"A munkalap nem került az első helyre. Várt: 'Lap 2', Kapott: '{workbook.Worksheet(1).Name}'");
                }
            });

            runTest("SetSheetPosition: Túl alacsony pozíciót (0) az első helyre korrigál", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    workbook.Worksheets.Add("Lap 1");
                    var sheetToMove = workbook.Worksheets.Add("Lap 2");

                    // Act
                    WorksheetExtensions.SetSheetPosition(sheetToMove, 0);

                    // Assert
                    if (sheetToMove.Position != 1)
                        throw new Exception($"A pozíció nem lett 1-re korrigálva. Várt: 1, Kapott: {sheetToMove.Position}");
                }
            });

            runTest("SetSheetPosition: Túl magas pozíciót az utolsó helyre korrigál", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var sheetToMove = workbook.Worksheets.Add("Lap 1");
                    workbook.Worksheets.Add("Lap 2");
                    
                    // Act
                    WorksheetExtensions.SetSheetPosition(sheetToMove, 99);

                    // Assert
                    if (sheetToMove.Position != 2)
                        throw new Exception($"A pozíció nem lett az utolsó helyre (2) korrigálva. Várt: 2, Kapott: {sheetToMove.Position}");
                }
            });

            runTest("SetSheetPosition: Nem változtat, ha a pozíció a jelenlegi", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    workbook.Worksheets.Add("Lap 1");
                    var sheetToMove = workbook.Worksheets.Add("Lap 2");
                    var initialOrder = workbook.Worksheets.Select(ws => ws.Name).ToList();

                    // Act
                    WorksheetExtensions.SetSheetPosition(sheetToMove, 2);

                    // Assert
                    var finalOrder = workbook.Worksheets.Select(ws => ws.Name).ToList();
                    if (!initialOrder.SequenceEqual(finalOrder))
                        throw new Exception("A munkalapok sorrendje hibásan megváltozott.");
                }
            });

            runTest("SetSheetPosition: Null worksheet esetén ArgumentNullException-t dob", () =>
            {
                // Arrange
                IXLWorksheet nullSheet = null;

                // Act & Assert
                try
                {
                    WorksheetExtensions.SetSheetPosition(nullSheet, 1);
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
