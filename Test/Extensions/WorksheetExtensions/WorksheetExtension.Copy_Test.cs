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
    public class WorksheetExtensions_Copy_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorksheetExtCopyTests", Guid.NewGuid().ToString());
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

            runTest("Copy: Sikeresen lemásol egy munkalapot tartalommal és stílussal", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var originalWs = workbook.Worksheets.Add("Eredeti");
                    var cell = originalWs.Cell("A1");
                    cell.Value = "Másolandó adat";
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.Yellow;

                    // Act
                    var newWs = WorksheetExtensions.Copy(originalWs, "Másolat");

                    // Assert
                    if (!workbook.Worksheets.Contains("Másolat"))
                        throw new Exception("A másolt munkalap nem jött létre.");
                    
                    var copiedCell = newWs.Cell("A1");
                    if (copiedCell.GetValue<string>() != "Másolandó adat")
                        throw new Exception($"A tartalom nem másolódott át. Várt: 'Másolandó adat', Kapott: '{copiedCell.Value}'");
                    if (!copiedCell.Style.Font.Bold)
                        throw new Exception("A 'Bold' stílus nem másolódott át.");
                    if (copiedCell.Style.Fill.BackgroundColor.Color != XLColor.Yellow.Color)
                        throw new Exception("A háttérszín nem másolódott át.");
                }
            });

            runTest("Copy: Üres munkalapot is helyesen másol", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var originalWs = workbook.Worksheets.Add("Üres Eredeti");

                    // Act
                    var newWs = WorksheetExtensions.Copy(originalWs, "Üres Másolat");

                    // Assert
                    if (!workbook.Worksheets.Contains("Üres Másolat"))
                        throw new Exception("Az üres munkalap másolata nem jött létre.");
                    if (newWs.LastCellUsed() != null)
                        throw new Exception($"A másolt munkalap nem üres. Utolsó cella: {newWs.LastCellUsed().Address}");
                }
            });

            runTest("Copy: ArgumentException-t dob null vagy üres új névre", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("Eredeti");
                    
                    // Act & Assert for null
                    try
                    {
                        WorksheetExtensions.Copy(ws, null);
                        throw new Exception("Nem dobódott a várt ArgumentException null névre.");
                    }
                    catch (ArgumentException) { /* Test Passed */ }

                    // Act & Assert for whitespace
                    try
                    {
                        WorksheetExtensions.Copy(ws, "   ");
                        throw new Exception("Nem dobódott a várt ArgumentException üres névre.");
                    }
                    catch (ArgumentException) { /* Test Passed */ }
                }
            });

            runTest("Copy: Kivételt dob, ha a cél munkalapnév már létezik", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    workbook.Worksheets.Add("Eredeti");
                    workbook.Worksheets.Add("Már Létezik");

                    // Act & Assert
                    try
                    {
                        WorksheetExtensions.Copy(workbook.Worksheet("Eredeti"), "Már Létezik");
                        throw new Exception("Nem dobódott kivétel, amikor egy létező nevű munkalapra próbáltunk másolni.");
                    }
                    catch (ArgumentException ex)
                    {
                        // JAVÍTVA: Az ellenőrzés most már kis- és nagybetűre érzéketlen.
                        if (!ex.Message.Contains("already been added", StringComparison.OrdinalIgnoreCase))
                            throw new Exception($"Váratlan hibaüzenet: {ex.Message}");
                        /* Test Passed */
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
