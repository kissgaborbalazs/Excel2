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
    public class WorksheetExtensions_GetLastColumn_Row_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorksheetExtGetLastColRowTests", Guid.NewGuid().ToString());
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

            runTest("GetLastColumn(row): Helyesen adja vissza az utolsó oszlop betűjelét", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A5").Value = "Adat1";
                    ws.Cell("B5").Value = "Adat2";
                    ws.Cell("C5").Value = "Utolsó Adat"; // C oszlop

                    // Act
                    var lastColumnLetter = WorksheetExtensions.GetLastColumn(ws, 5);

                    // Assert
                    if (lastColumnLetter != "C")
                        throw new Exception($"Hibás oszlop betűjelét adta vissza. Várt: 'C', Kapott: '{lastColumnLetter}'");
                }
            });

            runTest("GetLastColumn(row): Üres sorra üres stringet ad vissza", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Adat egy másik sorban";

                    // Act
                    var lastColumnLetter = WorksheetExtensions.GetLastColumn(ws, 2); // A 2. sor üres

                    // Assert
                    if (lastColumnLetter != string.Empty)
                        throw new Exception($"Nem üres stringet adott vissza üres sorra. Várt: '', Kapott: '{lastColumnLetter}'");
                }
            });

            runTest("GetLastColumn(row): Helyesen kezeli a kihagyásokat a sorban", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A3").Value = "Első";
                    // B3 üres
                    ws.Cell("C3").Value = "Utolsó"; // C oszlop

                    // Act
                    var lastColumnLetter = WorksheetExtensions.GetLastColumn(ws, 3);

                    // Assert
                    if (lastColumnLetter != "C")
                        throw new Exception($"Hibásan kezelte a kihagyást. Várt: 'C', Kapott: '{lastColumnLetter}'");
                }
            });

            runTest("GetLastColumn(row): Érvénytelen sorszám esetén kivételt dob", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    
                    // Act & Assert
                    try
                    {
                        WorksheetExtensions.GetLastColumn(ws, 0);
                        throw new Exception("Nem dobódott a várt ArgumentOutOfRangeException érvénytelen sorszámra.");
                    }
                    catch (ArgumentOutOfRangeException)
                    {
                        // A ClosedXML helyesen dob kivételt, ha a sorszám < 1
                        /* Test Passed */
                    }
                }
            });

            runTest("GetLastColumn(row): Üres munkalapon üres stringet ad vissza", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");

                    // Act
                    var lastColumnLetter = WorksheetExtensions.GetLastColumn(ws, 1);

                    // Assert
                    if (lastColumnLetter != string.Empty)
                        throw new Exception($"Nem üres stringet adott vissza üres munkalapra. Várt: '', Kapott: '{lastColumnLetter}'");
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