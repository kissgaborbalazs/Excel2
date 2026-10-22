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

namespace MPRT.Excel.Tests.Activities
{
    public class AutoFitColumns_Activity_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "AutoFitActivityTests", Guid.NewGuid().ToString());
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

            runTest("AutoFitColumns Activity: Csak a megadott oszlopokat igazítja", () =>
            {
                // Arrange: Tesztfájl létrehozása
                var filePath = Path.Combine(testDir, "test_specific_cols.xlsx");
                var sheetName = "Adatok";
                double initialWidthA, initialWidthB, initialWidthC;

                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add(sheetName);
                    ws.Cell("A1").Value = "Rövid";
                    ws.Cell("B1").Value = "Ez egy sokkal hosszabb szöveg";
                    ws.Cell("C1").Value = "Ez marad";
                    initialWidthA = ws.Column(1).Width;
                    initialWidthB = ws.Column(2).Width;
                    initialWidthC = ws.Column(3).Width;
                    workbook.SaveAs(filePath);
                }

                // Act: Activity meghívása
                var activity = new AutoFitColumns();
                var inputs = new Dictionary<string, object>
                {
                    { "FilePath", filePath },
                    { "SheetName", sheetName },
                    { "Columns", new int[] { 1, 2 } } // Csak az A és B oszlopot igazítjuk
                };
                WorkflowInvoker.Invoke(activity, inputs);

                // Assert: Eredmény validálása
                using (var resultWb = new XLWorkbook(filePath))
                {
                    var resultWs = resultWb.Worksheet(sheetName);
                    var finalWidthA = resultWs.Column(1).Width;
                    var finalWidthB = resultWs.Column(2).Width;
                    var finalWidthC = resultWs.Column(3).Width;

                    if (finalWidthA >= initialWidthA)
                        throw new Exception($"Az 'A' oszlop szélessége nem csökkent. Várt: >{initialWidthA}, Kapott: {finalWidthA}");
                    if (finalWidthB <= initialWidthB)
                        throw new Exception($"A 'B' oszlop szélessége nem nőtt. Várt: >{initialWidthB}, Kapott: {finalWidthB}");
                    if (finalWidthC != initialWidthC)
                        throw new Exception($"A 'C' oszlop szélessége hibásan megváltozott. Várt: {initialWidthC}, Kapott: {finalWidthC}");
                }
            });
            
            runTest("AutoFitColumns Activity: Az összes oszlopot igazítja, ha nincs megadva paraméter", () =>
            {
                // Arrange
                var filePath = Path.Combine(testDir, "test_all_cols.xlsx");
                var sheetName = "MindenOszlop";
                double initialWidthA, initialWidthB;

                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add(sheetName);
                    ws.Cell("A1").Value = "Széles tartalom az A oszlopban";
                    ws.Cell("B1").Value = "Ez is széles";
                    initialWidthA = ws.Column(1).Width;
                    initialWidthB = ws.Column(2).Width;
                    workbook.SaveAs(filePath);
                }

                // Act
                var activity = new AutoFitColumns();
                var inputs = new Dictionary<string, object>
                {
                    { "FilePath", filePath },
                    { "SheetName", sheetName },
                    { "Columns", null } // Nem adunk meg oszlopokat
                };
                WorkflowInvoker.Invoke(activity, inputs);

                // Assert
                using (var resultWb = new XLWorkbook(filePath))
                {
                    var resultWs = resultWb.Worksheet(sheetName);
                    var finalWidthA = resultWs.Column(1).Width;
                    var finalWidthB = resultWs.Column(2).Width;

                    if (finalWidthA <= initialWidthA)
                        throw new Exception($"Az 'A' oszlop szélessége nem nőtt. Várt: >{initialWidthA}, Kapott: {finalWidthA}");
                    if (finalWidthB <= initialWidthB)
                        throw new Exception($"A 'B' oszlop szélessége nem nőtt. Várt: >{initialWidthB}, Kapott: {finalWidthB}");
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