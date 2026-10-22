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
    public class RangeExtensions_SetNumberFormat_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "RangeExtSetNumFormatTests", Guid.NewGuid().ToString());
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

            runTest("SetNumberFormat(enum): Helyesen beállítja a Dátum formátumot", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    var range = ws.Range("A1:A5");
                    range.Value = DateTime.Now;

                    // Act
                    RangeExtensions.SetNumberFormat(range, NumberFormatType.Date);

                    // Assert
                    var expectedFormat = "yyyy.mm.dd";
                    if (range.Style.NumberFormat.Format != expectedFormat)
                        throw new Exception($"A számformátum hibás. Várt: '{expectedFormat}', Kapott: '{range.Style.NumberFormat.Format}'");
                }
            });

            runTest("SetNumberFormat(enum): Helyesen beállítja a Pénznem formátumot", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    var range = ws.Range("B1");
                    range.Value = 12345.67;

                    // Act
                    RangeExtensions.SetNumberFormat(range, NumberFormatType.Currency);

                    // Assert
                    var expectedFormat = "#,##0.00 [$Ft-40E]";
                    if (range.Style.NumberFormat.Format != expectedFormat)
                        throw new Exception($"A számformátum hibás. Várt: '{expectedFormat}', Kapott: '{range.Style.NumberFormat.Format}'");
                }
            });

            runTest("SetNumberFormat(string): Helyesen beállít egy egyedi számformátumot", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    var range = ws.Range("C1");
                    range.Value = 1234;
                    var customFormat = "00000"; // Öt számjegyű, vezető nullákkal

                    // Act
                    RangeExtensions.SetNumberFormat(range, customFormat);

                    // Assert
                    if (range.Style.NumberFormat.Format != customFormat)
                        throw new Exception($"A számformátum hibás. Várt: '{customFormat}', Kapott: '{range.Style.NumberFormat.Format}'");
                    if (range.Cell(1,1).GetFormattedString() != "01234")
                        throw new Exception($"A formázott érték hibás. Várt: '01234', Kapott: '{range.Cell(1,1).GetFormattedString()}'");
                }
            });

            runTest("SetNumberFormat(string): Üres vagy null formátummal nem változtat", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    var range = ws.Range("D1");
                    range.Value = 100;
                    range.Style.NumberFormat.Format = "General"; // Kezdeti állapot
                    var initialFormat = range.Style.NumberFormat.Format;

                    // Act
                    RangeExtensions.SetNumberFormat(range, "");
                    RangeExtensions.SetNumberFormat(range, "   ");
                    RangeExtensions.SetNumberFormat(range, (string)null);

                    // Assert
                    if (range.Style.NumberFormat.Format != initialFormat)
                        throw new Exception($"A formátum hibásan megváltozott. Várt: '{initialFormat}', Kapott: '{range.Style.NumberFormat.Format}'");
                }
            });

            runTest("SetNumberFormat: Null range esetén ArgumentNullException-t dob", () =>
            {
                // Arrange
                IXLRange nullRange = null;

                // Act & Assert (enum)
                try
                {
                    RangeExtensions.SetNumberFormat(nullRange, NumberFormatType.General);
                    throw new Exception("Nem dobódott a várt ArgumentNullException (enum).");
                }
                catch (ArgumentNullException) { /* Test Passed */ }

                // Act & Assert (string)
                try
                {
                    RangeExtensions.SetNumberFormat(nullRange, "General");
                    throw new Exception("Nem dobódott a várt ArgumentNullException (string).");
                }
                catch (ArgumentNullException) { /* Test Passed */ }
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
