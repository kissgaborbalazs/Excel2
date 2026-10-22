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
    public class WorksheetExtensions_MoveColumn_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorksheetExtMoveColTests", Guid.NewGuid().ToString());
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

            runTest("MoveColumn: Oszlopot jobbra mozgat (A -> C)", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "A Oszlop";
                    ws.Cell("B1").Value = "B Oszlop";
                    ws.Cell("C1").Value = "C Oszlop";
                    ws.Column(1).Style.Fill.BackgroundColor = XLColor.Yellow;

                    // Act
                    WorksheetExtensions.MoveColumn(ws, 1, 3);

                    // Assert
                    if (ws.Cell("A1").GetValue<string>() != "B Oszlop")
                        throw new Exception($"Az 'A' oszlop tartalma hibás. Várt: 'B Oszlop', Kapott: '{ws.Cell("A1").Value}'");
                    if (ws.Cell("B1").GetValue<string>() != "C Oszlop")
                        throw new Exception($"A 'B' oszlop tartalma hibás. Várt: 'C Oszlop', Kapott: '{ws.Cell("B1").Value}'");
                    if (ws.Cell("C1").GetValue<string>() != "A Oszlop")
                        throw new Exception($"A 'C' oszlop tartalma hibás. Várt: 'A Oszlop', Kapott: '{ws.Cell("C1").Value}'");
                    if (ws.Column(3).Style.Fill.BackgroundColor.Color != XLColor.Yellow.Color)
                        throw new Exception("A stílus nem lett áthelyezve.");
                }
            });

            runTest("MoveColumn: Oszlopot balra mozgat (C -> A)", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "A Oszlop";
                    ws.Cell("B1").Value = "B Oszlop";
                    ws.Cell("C1").Value = "C Oszlop";
                    ws.Column(3).Style.Font.Bold = true;

                    // Act
                    WorksheetExtensions.MoveColumn(ws, 3, 1);

                    // Assert
                    if (ws.Cell("A1").GetValue<string>() != "C Oszlop")
                        throw new Exception($"Az 'A' oszlop tartalma hibás. Várt: 'C Oszlop', Kapott: '{ws.Cell("A1").Value}'");
                    if (ws.Cell("B1").GetValue<string>() != "A Oszlop")
                        throw new Exception($"A 'B' oszlop tartalma hibás. Várt: 'A Oszlop', Kapott: '{ws.Cell("B1").Value}'");
                    if (ws.Cell("C1").GetValue<string>() != "B Oszlop")
                        throw new Exception($"A 'C' oszlop tartalma hibás. Várt: 'B Oszlop', Kapott: '{ws.Cell("C1").Value}'");
                    if (!ws.Column(1).Style.Font.Bold)
                        throw new Exception("A stílus nem lett áthelyezve.");
                }
            });

            runTest("MoveColumn: Érvénytelen paraméterekkel nem csinál semmit", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "A"; ws.Cell("B1").Value = "B";
                    
                    Action<int, int> checkNoChange = (source, dest) =>
                    {
                        WorksheetExtensions.MoveColumn(ws, source, dest);
                        if (ws.Cell("A1").GetValue<string>() != "A" || ws.Cell("B1").GetValue<string>() != "B")
                            throw new Exception($"A munkalap hibásan megváltozott a ({source}, {dest}) hívásra.");
                    };

                    // Act & Assert
                    checkNoChange(0, 2);
                    checkNoChange(2, 0);
                    checkNoChange(-1, 2);
                    checkNoChange(2, -1);
                    checkNoChange(2, 2);
                }
            });

            runTest("MoveColumn: Üres munkalapon nem okoz hibát", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");

                    // Act & Assert (nem dobhat kivételt)
                    WorksheetExtensions.MoveColumn(ws, 1, 2);
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