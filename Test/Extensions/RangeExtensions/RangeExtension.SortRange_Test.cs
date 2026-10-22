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
    public class RangeExtensions_SortRange_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "RangeExtSortRangeTests", Guid.NewGuid().ToString());
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
            // TESZTESETEK
            // =================================================================

            runTest("SortRange(int): Növekvő rendezés fejléccel", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Név"; ws.Cell("B1").Value = "Pont";
                    ws.Cell("A2").Value = "Béla"; ws.Cell("B2").Value = 80;
                    ws.Cell("A3").Value = "Anna"; ws.Cell("B3").Value = 95;
                    ws.Cell("A4").Value = "Cecil"; ws.Cell("B4").Value = 70;
                    var range = ws.Range("A1:B4");

                    // Act
                    RangeExtensions.SortRange(range, 2, XLSortOrder.Ascending, true);

                    // Assert
                    if (ws.Cell("A1").GetValue<string>() != "Név") throw new Exception("A fejléc elmozdult.");
                    var expected = new List<int> { 70, 80, 95 };
                    var actual = ws.Range("B2:B4").Cells().Select(c => c.GetValue<int>()).ToList();
                    if (!expected.SequenceEqual(actual)) throw new Exception($"A sorrend hibás. Várt: {string.Join(",", expected)}, Kapott: {string.Join(",", actual)}");
                }
            });

            runTest("SortRange(int): Csökkenő rendezés fejléc nélkül", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = 80;
                    ws.Cell("A2").Value = 95;
                    ws.Cell("A3").Value = 70;
                    var range = ws.Range("A1:A3");

                    // Act
                    RangeExtensions.SortRange(range, 1, XLSortOrder.Descending, false);

                    // Assert
                    var expected = new List<int> { 95, 80, 70 };
                    var actual = range.Cells().Select(c => c.GetValue<int>()).ToList();
                    if (!expected.SequenceEqual(actual)) throw new Exception($"A sorrend hibás. Várt: {string.Join(",", expected)}, Kapott: {string.Join(",", actual)}");
                }
            });

            runTest("SortRange(string): Oszlopbetűjel alapján rendez", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("B2").Value = "ID"; ws.Cell("C2").Value = "Név";
                    ws.Cell("B3").Value = 3; ws.Cell("C3").Value = "Cecil";
                    ws.Cell("B4").Value = 1; ws.Cell("C4").Value = "Anna";
                    var range = ws.Range("B2:C4");

                    // Act
                    RangeExtensions.SortRange(range, "B", XLSortOrder.Ascending, true);

                    // Assert
                    var expected = new List<string> { "Anna", "Cecil" };
                    var actual = ws.Range("C3:C4").Cells().Select(c => c.GetValue<string>()).ToList();
                    if (!expected.SequenceEqual(actual)) throw new Exception($"A sorrend hibás. Várt: {string.Join(",", expected)}, Kapott: {string.Join(",", actual)}");
                }
            });

            runTest("SortRangeByName: Oszlopnév alapján rendez", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Név"; ws.Cell("B1").Value = "Pont";
                    ws.Cell("A2").Value = "Béla"; ws.Cell("B2").Value = 80;
                    ws.Cell("A3").Value = "Anna"; ws.Cell("B3").Value = 95;
                    var range = ws.Range("A1:B3");

                    // Act
                    RangeExtensions.SortRangeByName(range, "Pont", XLSortOrder.Descending);

                    // Assert
                    if (ws.Cell("A2").GetValue<string>() != "Anna") throw new Exception($"A sorrend hibás. Várt: 'Anna', Kapott: '{ws.Cell("A2").GetValue<string>()}'");
                }
            });

            runTest("SortRangeByName: Nem létező oszlopnév ArgumentException-t dob", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Név";
                    var range = ws.Range("A1:A2");
                    
                    // Act & Assert
                    try
                    {
                        RangeExtensions.SortRangeByName(range, "Hibás Név");
                        throw new Exception("Nem dobódott a várt ArgumentException.");
                    }
                    catch (ArgumentException) { /* Test Passed */ }
                }
            });
            
            runTest("SortRange(int): Túl nagy oszlop sorszám ArgumentOutOfRangeException-t dob", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    var range = ws.Range("A1:B2");
                    
                    // Act & Assert
                    try
                    {
                        RangeExtensions.SortRange(range, 99);
                        throw new Exception("Nem dobódott a várt ArgumentOutOfRangeException.");
                    }
                    catch (ArgumentOutOfRangeException) { /* Test Passed */ }
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