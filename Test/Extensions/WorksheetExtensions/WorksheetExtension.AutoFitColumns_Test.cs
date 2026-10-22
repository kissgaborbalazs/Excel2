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
    public class WorksheetExtensions_AutoFitColumns_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorksheetExtAutoFitTests", Guid.NewGuid().ToString());
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

            runTest("AutoFitColumns: Paraméterek nélkül az összes oszlopot igazítja", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Rövid";
                    ws.Cell("B1").Value = "Ez egy sokkal hosszabb szöveg";
                    ws.Cell("C1").Value = "Normál";
                    
                    var initialWidthA = ws.Column(1).Width;
                    var initialWidthB = ws.Column(2).Width;

                    // Act
                    WorksheetExtensions.AutoFitColumns(ws);

                    // Assert
                    var finalWidthA = ws.Column(1).Width;
                    var finalWidthB = ws.Column(2).Width;

                    if (finalWidthA >= initialWidthA)
                        throw new Exception($"Az 'A' oszlop szélessége nem csökkent. Várt: <{initialWidthA}, Kapott: {finalWidthA}");
                    if (finalWidthB <= initialWidthB)
                        throw new Exception($"A 'B' oszlop szélessége nem nőtt. Várt: >{initialWidthB}, Kapott: {finalWidthB}");
                }
            });

            runTest("AutoFitColumns: Csak a megadott oszlopokat igazítja", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Ezt igazítani kell";
                    ws.Cell("B1").Value = "Ez maradjon változatlan";
                    ws.Cell("C1").Value = "Ezt is igazítani kell";

                    var initialWidthA = ws.Column(1).Width;
                    var initialWidthB = ws.Column(2).Width;
                    var initialWidthC = ws.Column(3).Width;

                    // Act
                    WorksheetExtensions.AutoFitColumns(ws, 1, 3);

                    // Assert
                    var finalWidthA = ws.Column(1).Width;
                    var finalWidthB = ws.Column(2).Width;
                    var finalWidthC = ws.Column(3).Width;

                    if (finalWidthA <= initialWidthA)
                        throw new Exception($"Az 'A' oszlop szélessége nem változott. Várt: >{initialWidthA}, Kapott: {finalWidthA}");
                    if (finalWidthB != initialWidthB)
                        throw new Exception($"A 'B' oszlop szélessége hibásan megváltozott. Várt: {initialWidthB}, Kapott: {finalWidthB}");
                    if (finalWidthC <= initialWidthC)
                        throw new Exception($"A 'C' oszlop szélessége nem változott. Várt: >{initialWidthC}, Kapott: {finalWidthC}");
                }
            });

            runTest("AutoFitColumns: Üres munkalapon nem okoz hibát", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");

                    // Act & Assert (nem dobhat kivételt)
                    WorksheetExtensions.AutoFitColumns(ws);
                    WorksheetExtensions.AutoFitColumns(ws, 1, 2);
                }
            });

            runTest("AutoFitColumns: Érvénytelen (nulla vagy negatív) oszlopindexet figyelmen kívül hagy", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Hosszú szöveg";
                    var initialWidthA = ws.Column(1).Width;

                    // Act & Assert (nem dobhat kivételt)
                    WorksheetExtensions.AutoFitColumns(ws, 0, -5, 1);
                    
                    var finalWidthA = ws.Column(1).Width;
                    if (finalWidthA <= initialWidthA)
                        throw new Exception("Az érvényes oszlop (1) nem lett igazítva az érvénytelenek mellett.");
                }
            });

            runTest("AutoFitColumns: Null oszlop-tömb esetén az összes oszlopot igazítja", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Széles tartalom";
                    var initialWidthA = ws.Column(1).Width;

                    // Act
                    WorksheetExtensions.AutoFitColumns(ws, null);

                    // Assert
                    var finalWidthA = ws.Column(1).Width;
                    if (finalWidthA <= initialWidthA)
                        throw new Exception($"Az oszlop szélessége nem nőtt. Várt: >{initialWidthA}, Kapott: {finalWidthA}");
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