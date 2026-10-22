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
    public class WorksheetExtensions_SetFormula_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorksheetExtSetFormulaTests", Guid.NewGuid().ToString());
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

            runTest("SetFormula: Helyesen beállít egy érvényes képletet", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    var formula = "=SUM(A1:B1)";

                    // Act
                    WorksheetExtensions.SetFormula(ws, "C1", formula);

                    // Assert
                    var cell = ws.Cell("C1");
                    if (!cell.HasFormula)
                        throw new Exception("A cella nem tartalmaz képletet a beállítás után.");
                    if (cell.FormulaA1 != formula.Replace("=", ""))
                        throw new Exception($"A beállított képlet hibás. Várt: '{formula}', Kapott: '{cell.FormulaA1}'");
                }
            });

            runTest("SetFormula: Felülír egy meglévő értéket", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    var cell = ws.Cell("A1");
                    cell.Value = "Régi Érték";

                    // Act
                    WorksheetExtensions.SetFormula(ws, "A1", "=1+1");

                    // Assert
                    if (!cell.HasFormula)
                        throw new Exception("A képlet nem írta felül a meglévő értéket.");
                    if (cell.GetValue<double>() != 2)
                        throw new Exception("A képlet kiértékelt értéke hibás.");
                }
            });

            runTest("SetFormula: Nem csinál semmit null vagy üres paraméterekkel", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Eredeti";

                    // Act
                    WorksheetExtensions.SetFormula(ws, null, "=1+1");
                    WorksheetExtensions.SetFormula(ws, "   ", "=1+1");
                    WorksheetExtensions.SetFormula(ws, "B1", null);
                    WorksheetExtensions.SetFormula(ws, "B1", "   ");

                    // Assert
                    if (ws.Cell("A1").GetValue<string>() != "Eredeti")
                        throw new Exception("A cella értéke hibásan megváltozott érvénytelen címre.");
                    if (ws.Cell("B1").HasFormula || !ws.Cell("B1").IsEmpty())
                        throw new Exception("A B1 cella hibásan kapott képletet érvénytelen formula stringgel.");
                }
            });

            runTest("SetFormula: Null worksheet esetén ArgumentNullException-t dob", () =>
            {
                // Arrange
                IXLWorksheet nullSheet = null;

                // Act & Assert
                try
                {
                    WorksheetExtensions.SetFormula(nullSheet, "A1", "=1");
                    throw new Exception("Nem dobódott a várt ArgumentNullException.");
                }
                catch (ArgumentNullException)
                {
                    /* Test Passed */
                }
            });

            runTest("SetFormula: Érvénytelen cellacím esetén kivételt dob", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");

                    // Act & Assert
                    try
                    {
                        WorksheetExtensions.SetFormula(ws, "ÉrvénytelenCím", "=1");
                        throw new Exception("Nem dobódott a várt kivétel érvénytelen címre.");
                    }
                    catch (ArgumentException)
                    {
                        // A ClosedXML ArgumentException-t dob érvénytelen címre
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
