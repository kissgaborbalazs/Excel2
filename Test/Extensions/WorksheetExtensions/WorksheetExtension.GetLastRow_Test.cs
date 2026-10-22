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
    public class WorksheetExtensions_GetLastRow_Col_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorksheetExtGetLastRowColTests", Guid.NewGuid().ToString());
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

            runTest("GetLastRow(col): Helyesen adja vissza az utolsó sor számát egy adott oszlopban", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("B1").Value = "Fejléc";
                    ws.Cell("B5").Value = "Adat1";
                    ws.Cell("B10").Value = "Utolsó Adat"; // 10. sor
                    ws.Cell("C12").Value = "Másik oszlop";

                    // Act
                    var lastRow = WorksheetExtensions.GetLastRow(ws, "B");

                    // Assert
                    if (lastRow != 10)
                        throw new Exception($"Hibás sorszámot adott vissza. Várt: 10, Kapott: {lastRow}");
                }
            });

            runTest("GetLastRow(col): 0-t ad vissza üres oszlopra", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Adat egy másik oszlopban";

                    // Act
                    var lastRow = WorksheetExtensions.GetLastRow(ws, "B"); // A B oszlop üres

                    // Assert
                    if (lastRow != 0)
                        throw new Exception($"Nem 0-t adott vissza üres oszlopra. Várt: 0, Kapott: {lastRow}");
                }
            });

            runTest("GetLastRow(col): Helyesen kezeli a kihagyásokat az oszlopban", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("C1").Value = "Első";
                    // C2, C3 üres
                    ws.Cell("C4").Value = "Utolsó"; // 4. sor

                    // Act
                    var lastRow = WorksheetExtensions.GetLastRow(ws, "C");

                    // Assert
                    if (lastRow != 4)
                        throw new Exception($"Hibásan kezelte a kihagyást. Várt: 4, Kapott: {lastRow}");
                }
            });

            runTest("GetLastRow(col): Érvénytelen oszlopbetűjel esetén 0-t ad vissza", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Adat";
                    
                    // Act
                    var lastRow = WorksheetExtensions.GetLastRow(ws, "123");

                    // Assert
                    if (lastRow != 0)
                    {
                        throw new Exception($"Hibás sorszámot adott vissza érvénytelen oszlopnévre. Várt: 0, Kapott: {lastRow}");
                    }
                }
            });

            runTest("GetLastRow(col): Üres munkalapon 0-t ad vissza", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");

                    // Act
                    var lastRow = WorksheetExtensions.GetLastRow(ws, "A");

                    // Assert
                    if (lastRow != 0)
                        throw new Exception($"Nem 0-t adott vissza üres munkalapra. Várt: 0, Kapott: {lastRow}");
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