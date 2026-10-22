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
    public class WorksheetExtensions_Delete_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorksheetExtDeleteTests", Guid.NewGuid().ToString());
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

            runTest("Delete: Sikeresen töröl egy munkalapot", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    workbook.Worksheets.Add("Törlendő Lap");
                    workbook.Worksheets.Add("Megmaradó Lap");
                    var sheetToDelete = workbook.Worksheet("Törlendő Lap");

                    // Act
                    WorksheetExtensions.Delete(sheetToDelete);

                    // Assert
                    if (workbook.Worksheets.Contains("Törlendő Lap"))
                        throw new Exception("A munkalap nem lett törölve a kollekcióból.");
                    if (workbook.Worksheets.Count != 1)
                        throw new Exception($"A munkalapok száma hibás. Várt: 1, Kapott: {workbook.Worksheets.Count}");
                    if (workbook.Worksheets.First().Name != "Megmaradó Lap")
                        throw new Exception("A rossz munkalap lett törölve.");
                }
            });

            runTest("Delete: A törölt munkalap objektum érvénytelenné válik", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var sheetToDelete = workbook.Worksheets.Add("Eltűnő Lap");

                    // Act
                    WorksheetExtensions.Delete(sheetToDelete);

                    // Assert
                    try
                    {
                        // JAVÍTVA: Egy összetettebb műveletet, például cella elérést
                        // próbálunk végrehajtani, ami megbízhatóbban kiváltja a hibát.
                        var cell = sheetToDelete.Cell("A1");
                        throw new Exception("Nem dobódott kivétel egy törölt munkalap használatakor.");
                    }
                    catch (InvalidOperationException)
                    {
                        // A ClosedXML InvalidOperationException-t dob, ha egy törölt lapot próbálunk használni.
                        /* Test Passed */
                    }
                }
            });

            runTest("Delete: Törlés nem hatással más munkalapok tartalmára", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    workbook.Worksheets.Add("Lap 1");
                    var sheetToDelete = workbook.Worksheets.Add("Lap 2 (Törlendő)");
                    var sheetToKeep = workbook.Worksheets.Add("Lap 3");
                    sheetToKeep.Cell("A1").Value = "Fontos Adat";

                    // Act
                    WorksheetExtensions.Delete(sheetToDelete);

                    // Assert
                    var keptSheet = workbook.Worksheet("Lap 3");
                    if (keptSheet.Cell("A1").GetValue<string>() != "Fontos Adat")
                    {
                        throw new Exception("Egy másik munkalap törlése megváltoztatta egy megmaradó lap tartalmát.");
                    }
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
