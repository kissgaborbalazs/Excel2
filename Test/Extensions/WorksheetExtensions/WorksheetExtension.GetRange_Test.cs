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
    public class WorksheetExtensions_GetRange_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorksheetExtGetRangeTests", Guid.NewGuid().ToString());
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

            runTest("GetRange: Kezdőcella nélkül a teljes használt területet adja vissza", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("B2").Value = "Start";
                    ws.Cell("D5").Value = "End"; // A használt terület B2:D5 lesz.

                    // Act
                    var range = WorksheetExtensions.GetRange(ws, null);

                    // Assert
                    var expectedAddress = "B2:D5";
                    if (range.RangeAddress.ToString() != expectedAddress)
                        throw new Exception($"Hibás tartományt adott vissza. Várt: '{expectedAddress}', Kapott: '{range.RangeAddress}'");
                }
            });

            runTest("GetRange: Kezdőcellából automatikusan felismer egy összefüggő táblázatot", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("B2").Value = "Fejléc1"; ws.Cell("C2").Value = "Fejléc2";
                    ws.Cell("B3").Value = "a";       ws.Cell("C3").Value = "b";
                    ws.Cell("B4").Value = "c";       ws.Cell("C4").Value = "d";

                    // Act
                    var range = WorksheetExtensions.GetRange(ws, "B2");

                    // Assert
                    var expectedAddress = "B2:C4";
                    if (range.RangeAddress.ToString() != expectedAddress)
                        throw new Exception($"Hibás tartományt ismert fel. Várt: '{expectedAddress}', Kapott: '{range.RangeAddress}'");
                }
            });

            runTest("GetRange: Helyesen kezeli a 'szaggatott' (jagged) táblázatokat", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "ID";    ws.Cell("B1").Value = "Név";  ws.Cell("C1").Value = "Megjegyzés";
                    ws.Cell("A2").Value = 1;       ws.Cell("B2").Value = "Anna";
                    ws.Cell("A3").Value = 2;       ws.Cell("B3").Value = "Béla"; ws.Cell("C3").Value = "Fontos"; // Ez a legalsó sor
                    ws.Cell("A4").Value = 3;       ws.Cell("B4").Value = "Cecil";

                    // Act
                    var range = WorksheetExtensions.GetRange(ws, "A1");

                    // Assert
                    var expectedAddress = "A1:C4";
                    if (range.RangeAddress.ToString() != expectedAddress)
                        throw new Exception($"Hibás tartományt ismert fel. Várt: '{expectedAddress}', Kapott: '{range.RangeAddress}'");
                }
            });

            runTest("GetRange: Csak egyetlen cellát ad vissza, ha a kezdőcella izolált", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("C3").Value = "Egyedüli cella";

                    // Act
                    var range = WorksheetExtensions.GetRange(ws, "C3");

                    // Assert
                    var expectedAddress = "C3:C3";
                    if (range.RangeAddress.ToString() != expectedAddress)
                        throw new Exception($"Hibás tartományt adott vissza. Várt: '{expectedAddress}', Kapott: '{range.RangeAddress}'");
                }
            });

            runTest("GetRange: Null worksheet esetén ArgumentNullException-t dob", () =>
            {
                // Arrange
                IXLWorksheet nullSheet = null;

                // Act & Assert
                try
                {
                    WorksheetExtensions.GetRange(nullSheet, "A1");
                    throw new Exception("Nem dobódott a várt ArgumentNullException.");
                }
                catch (ArgumentNullException)
                {
                    /* Test Passed */
                }
            });

            runTest("GetRange: Érvénytelen kezdőcella esetén kivételt dob", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    
                    // Act & Assert
                    try
                    {
                        WorksheetExtensions.GetRange(ws, "Érvénytelen");
                        throw new Exception("Nem dobódott a várt kivétel érvénytelen címre.");
                    }
                    catch (ArgumentException)
                    {
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