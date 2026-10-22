using System;
using System.IO;
using UiPath.CodedWorkflows;
using ClosedXML.Excel;
using System.Diagnostics;
using System.Linq;
using System.Collections.Generic;
using System.Activities;

// A tesztelendő activity névtere.
using MPRT.Excel.Activities.Format;
// A háttérben használt logikai könyvtár névtere.
using MPRT.Excel;


namespace MPRT.Excel.Tests.Activities
{
    public class FreezePanes_Activity_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "FreezePanesActivityTests", Guid.NewGuid().ToString());
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

            runTest("FreezePanes Activity: Helyesen rögzíti az első sort", () =>
            {
                // Arrange: Tesztfájl létrehozása
                var filePath = Path.Combine(testDir, "test_freeze_row.xlsx");
                var sheetName = "Jelentés";
                using (var workbook = new XLWorkbook())
                {
                    workbook.Worksheets.Add(sheetName).Cell("A1").Value = "Fejléc";
                    workbook.SaveAs(filePath);
                }

                // Act: Activity meghívása
                var activity = new FreezePanes();
                var inputs = new Dictionary<string, object>
                {
                    { "FilePath", filePath },
                    { "SheetName", sheetName },
                    { "RowsToFreeze", 1 },
                    { "ColumnsToFreeze", 0 }
                };
                WorkflowInvoker.Invoke(activity, inputs);

                // Assert: Eredmény validálása
                using (var resultWb = new XLWorkbook(filePath))
                {
                    var resultWs = resultWb.Worksheet(sheetName);
                    if (resultWs.SheetView.SplitRow != 1)
                        throw new Exception($"A rögzített sorok száma hibás. Várt: 1, Kapott: {resultWs.SheetView.SplitRow}");
                    if (resultWs.SheetView.SplitColumn != 0)
                        throw new Exception($"A rögzített oszlopok száma hibás. Várt: 0, Kapott: {resultWs.SheetView.SplitColumn}");
                }
            });
            
            runTest("FreezePanes Activity: Rögzíti a sorokat és oszlopokat is", () =>
            {
                // Arrange
                var filePath = Path.Combine(testDir, "test_freeze_both.xlsx");
                var sheetName = "Adatbázis";
                using (var workbook = new XLWorkbook())
                {
                    workbook.Worksheets.Add(sheetName).Cell("A1").Value = "Adat";
                    workbook.SaveAs(filePath);
                }

                // Act
                var activity = new FreezePanes();
                var inputs = new Dictionary<string, object>
                {
                    { "FilePath", filePath },
                    { "SheetName", sheetName },
                    { "RowsToFreeze", 2 },
                    { "ColumnsToFreeze", 3 }
                };
                WorkflowInvoker.Invoke(activity, inputs);

                // Assert
                using (var resultWb = new XLWorkbook(filePath))
                {
                    var resultWs = resultWb.Worksheet(sheetName);
                    if (resultWs.SheetView.SplitRow != 2)
                        throw new Exception($"A rögzített sorok száma hibás. Várt: 2, Kapott: {resultWs.SheetView.SplitRow}");
                    if (resultWs.SheetView.SplitColumn != 3)
                        throw new Exception($"A rögzített oszlopok száma hibás. Várt: 3, Kapott: {resultWs.SheetView.SplitColumn}");
                }
            });

            runTest("FreezePanes Activity: Feloldja a rögzítést (0, 0) paraméterekkel", () =>
            {
                // Arrange
                var filePath = Path.Combine(testDir, "test_unfreeze.xlsx");
                var sheetName = "Feloldás";
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add(sheetName);
                    ws.SheetView.Freeze(5, 5); // Előzetes rögzítés
                    workbook.SaveAs(filePath);
                }

                // Act
                var activity = new FreezePanes();
                var inputs = new Dictionary<string, object>
                {
                    { "FilePath", filePath },
                    { "SheetName", sheetName },
                    { "RowsToFreeze", 0 },
                    { "ColumnsToFreeze", 0 }
                };
                WorkflowInvoker.Invoke(activity, inputs);

                // Assert
                using (var resultWb = new XLWorkbook(filePath))
                {
                    var resultWs = resultWb.Worksheet(sheetName);
                    if (resultWs.SheetView.SplitRow != 0)
                        throw new Exception($"A sorrögzítés nem lett feloldva. Várt: 0, Kapott: {resultWs.SheetView.SplitRow}");
                    if (resultWs.SheetView.SplitColumn != 0)
                        throw new Exception($"Az oszloprögzítés nem lett feloldva. Várt: 0, Kapott: {resultWs.SheetView.SplitColumn}");
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