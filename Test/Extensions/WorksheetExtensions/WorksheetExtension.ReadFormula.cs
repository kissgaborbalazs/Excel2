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
    public class WorksheetExtensions_ReadFormula_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorksheetExtReadFormulaTests", Guid.NewGuid().ToString());
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

            runTest("ReadFormula: Visszaadja a helyes képletet", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    var expectedFormula = "SUM(A1:B1)";
                    ws.Cell("C1").FormulaA1 = $"={expectedFormula}";

                    // Act
                    var actualFormula = WorksheetExtensions.ReadFormula(ws, "C1");

                    // Assert
                    if (actualFormula != expectedFormula)
                        throw new Exception($"Hibás képletet adott vissza. Várt: '{expectedFormula}', Kapott: '{actualFormula}'");
                }
            });

            runTest("ReadFormula: Üres stringet ad vissza, ha a cella nem tartalmaz képletet", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Nem képlet";

                    // Act
                    var result = WorksheetExtensions.ReadFormula(ws, "A1");

                    // Assert
                    if (result != string.Empty)
                        throw new Exception($"Nem üres stringet adott vissza. Várt: '', Kapott: '{result}'");
                }
            });

            runTest("ReadFormula: Üres stringet ad vissza üres cellára", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");

                    // Act
                    var result = WorksheetExtensions.ReadFormula(ws, "B2");

                    // Assert
                    if (result != string.Empty)
                        throw new Exception($"Nem üres stringet adott vissza üres cellára. Várt: '', Kapott: '{result}'");
                }
            });

            runTest("ReadFormula: Üres stringet ad vissza null vagy üres címre", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").FormulaA1 = "=1+1";

                    // Act
                    var resultNull = WorksheetExtensions.ReadFormula(ws, null);
                    var resultEmpty = WorksheetExtensions.ReadFormula(ws, "");
                    var resultWhitespace = WorksheetExtensions.ReadFormula(ws, "   ");

                    // Assert
                    if (resultNull != string.Empty)
                        throw new Exception($"Nem üres stringet adott vissza null címre. Várt: '', Kapott: '{resultNull}'");
                    if (resultEmpty != string.Empty)
                        throw new Exception($"Nem üres stringet adott vissza üres címre. Várt: '', Kapott: '{resultEmpty}'");
                    if (resultWhitespace != string.Empty)
                        throw new Exception($"Nem üres stringet adott vissza whitespace címre. Várt: '', Kapott: '{resultWhitespace}'");
                }
            });

            runTest("ReadFormula: Érvénytelen cellacím esetén kivételt dob", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");

                    // Act & Assert
                    try
                    {
                        WorksheetExtensions.ReadFormula(ws, "ÉrvénytelenCím");
                        throw new Exception("Nem dobódott a várt kivétel érvénytelen címre.");
                    }
                    catch (ArgumentException)
                    {
                        // A ClosedXML ArgumentException-t dob érvénytelen címre
                        /* Test Passed */
                    }
                }
            });

            runTest("ReadFormula: Null worksheet esetén ArgumentNullException-t dob", () =>
            {
                // Arrange
                IXLWorksheet nullSheet = null;

                // Act & Assert
                try
                {
                    WorksheetExtensions.ReadFormula(nullSheet, "A1");
                    throw new Exception("Nem dobódott a várt ArgumentNullException.");
                }
                catch (ArgumentNullException)
                {
                    /* Test Passed */
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
