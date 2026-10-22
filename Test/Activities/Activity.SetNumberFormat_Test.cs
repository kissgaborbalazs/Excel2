using System;
using System.IO;
using UiPath.CodedWorkflows;
using ClosedXML.Excel;
using System.Diagnostics;
using System.Linq;
using System.Collections.Generic;
using System.Activities;

// A tesztelendő activity és a háttérben használt logikai könyvtár névterei.
using MPRT.Excel.Activities.Format;
using MPRT.Excel;

namespace MPRT.Excel.Tests.Activities
{
    public class SetNumberFormat_Activity_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "SetNumberFormatActivityTests", Guid.NewGuid().ToString());
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

            runTest("SetNumberFormat Activity: Előre definiált enummal (Date) formáz", () =>
            {
                // Arrange: Tesztfájl létrehozása
                var filePath = Path.Combine(testDir, "test_enum_date.xlsx");
                var sheetName = "Dátumok";
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add(sheetName);
                    ws.Cell("A1").Value = 45678; // Excel serial date for 2025-01-21
                    workbook.SaveAs(filePath);
                }

                // Act: Activity meghívása
                var activity = new SetNumberFormat();
                // JAVÍTVA: A nem InArgument tulajdonságot közvetlenül az activity-n állítjuk be.
                activity.FormatType = NumberFormatType.Date;

                var inputs = new Dictionary<string, object>
                {
                    { "FilePath", filePath },
                    { "SheetName", sheetName },
                    { "Range", "A1" }
                    // A "FormatType" kikerült a szótárból.
                };
                WorkflowInvoker.Invoke(activity, inputs);

                // Assert: Eredmény validálása
                using (var resultWb = new XLWorkbook(filePath))
                {
                    var resultCell = resultWb.Worksheet(sheetName).Cell("A1");
                    var expectedFormat = "yyyy.mm.dd";
                    if (resultCell.Style.NumberFormat.Format != expectedFormat)
                        throw new Exception($"A számformátum hibás. Várt: '{expectedFormat}', Kapott: '{resultCell.Style.NumberFormat.Format}'");
                }
            });
            
            runTest("SetNumberFormat Activity: Egyedi formátummal ('#,##0') formáz", () =>
            {
                // Arrange
                var filePath = Path.Combine(testDir, "test_custom_number.xlsx");
                var sheetName = "Számok";
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add(sheetName);
                    ws.Cell("B2").Value = 12345.67;
                    workbook.SaveAs(filePath);
                }

                // Act
                var activity = new SetNumberFormat();
                var customFormatString = "#,##0";
                var inputs = new Dictionary<string, object>
                {
                    { "FilePath", filePath },
                    { "SheetName", sheetName },
                    { "Range", "B2" },
                    { "CustomFormat", customFormatString }
                };
                WorkflowInvoker.Invoke(activity, inputs);

                // Assert
                using (var resultWb = new XLWorkbook(filePath))
                {
                    var resultCell = resultWb.Worksheet(sheetName).Cell("B2");
                    if (resultCell.Style.NumberFormat.Format != customFormatString)
                        throw new Exception($"A számformátum hibás. Várt: '{customFormatString}', Kapott: '{resultCell.Style.NumberFormat.Format}'");
                }
            });

            runTest("SetNumberFormat Activity: Az egyedi formátum felülírja az enumot", () =>
            {
                // Arrange
                var filePath = Path.Combine(testDir, "test_override.xlsx");
                var sheetName = "Felülírás";
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add(sheetName);
                    ws.Cell("A1").Value = 0.25;
                    workbook.SaveAs(filePath);
                }

                // Act
                var activity = new SetNumberFormat();
                // JAVÍTVA: A nem InArgument tulajdonságot közvetlenül az activity-n állítjuk be.
                activity.FormatType = NumberFormatType.Percentage; // Ez 2 tizedesjegyet állítana be

                var customFormatString = "0.0000%"; // 4 tizedesjegyű százalék
                var inputs = new Dictionary<string, object>
                {
                    { "FilePath", filePath },
                    { "SheetName", sheetName },
                    { "Range", "A1" },
                    { "CustomFormat", customFormatString } // Ennek kell nyernie
                };
                WorkflowInvoker.Invoke(activity, inputs);

                // Assert
                using (var resultWb = new XLWorkbook(filePath))
                {
                    var resultCell = resultWb.Worksheet(sheetName).Cell("A1");
                    if (resultCell.Style.NumberFormat.Format != customFormatString)
                        throw new Exception($"A számformátum hibás. Várt: '{customFormatString}', Kapott: '{resultCell.Style.NumberFormat.Format}'");
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