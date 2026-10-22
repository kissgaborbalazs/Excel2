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
    public class WorksheetExtensions_Rename_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorksheetExtRenameTests", Guid.NewGuid().ToString());
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

            runTest("Rename: Sikeresen átnevez egy munkalapot", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("Régi Név");

                    // Act
                    WorksheetExtensions.Rename(ws, "Új Név");

                    // Assert
                    if (ws.Name != "Új Név")
                        throw new Exception($"A munkalap neve nem változott meg. Várt: 'Új Név', Kapott: '{ws.Name}'");
                    if (!workbook.Worksheets.Contains("Új Név") || workbook.Worksheets.Contains("Régi Név"))
                        throw new Exception("A munkalap-kollekció nem frissült helyesen.");
                }
            });

            runTest("Rename: ArgumentException-t dob null vagy üres új névre", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("Eredeti");
                    
                    // Act & Assert for null
                    try
                    {
                        WorksheetExtensions.Rename(ws, null);
                        throw new Exception("Nem dobódott a várt ArgumentException null névre.");
                    }
                    catch (ArgumentException) { /* Test Passed */ }

                    // Act & Assert for whitespace
                    try
                    {
                        WorksheetExtensions.Rename(ws, "   ");
                        throw new Exception("Nem dobódott a várt ArgumentException üres névre.");
                    }
                    catch (ArgumentException) { /* Test Passed */ }
                }
            });

            runTest("Rename: Kivételt dob, ha az új név már létezik", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var wsToRename = workbook.Worksheets.Add("Átnevezendő");
                    workbook.Worksheets.Add("Már Létezik");

                    // Act & Assert
                    try
                    {
                        WorksheetExtensions.Rename(wsToRename, "Már Létezik");
                        throw new Exception("Nem dobódott kivétel, amikor egy létező névre próbáltunk átnevezni.");
                    }
                    catch (ArgumentException ex)
                    {
                        if (!ex.Message.Contains("already been added", StringComparison.OrdinalIgnoreCase))
                            throw new Exception($"Váratlan hibaüzenet: {ex.Message}");
                        /* Test Passed */
                    }
                }
            });
            
            runTest("Rename: Átnevezés speciális karaktereket tartalmazó névre", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("Régi");
                    var newName = "Új Név (2025-ös verzió)";

                    // Act
                    WorksheetExtensions.Rename(ws, newName);

                    // Assert
                    if (ws.Name != newName)
                        throw new Exception($"A speciális karaktereket tartalmazó név hibás. Várt: '{newName}', Kapott: '{ws.Name}'");
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
