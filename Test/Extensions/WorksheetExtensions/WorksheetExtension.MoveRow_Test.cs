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
    public class WorksheetExtensions_MoveRow_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorksheetExtMoveRowTests", Guid.NewGuid().ToString());
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

            runTest("MoveRow: Sort lefelé mozgat (1 -> 3)", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "1. Sor";
                    ws.Cell("A2").Value = "2. Sor";
                    ws.Cell("A3").Value = "3. Sor";
                    ws.Row(1).Style.Fill.BackgroundColor = XLColor.Yellow;

                    // Act
                    WorksheetExtensions.MoveRow(ws, 1, 3);

                    // Assert
                    if (ws.Cell("A1").GetValue<string>() != "2. Sor")
                        throw new Exception($"Az 1. sor tartalma hibás. Várt: '2. Sor', Kapott: '{ws.Cell("A1").Value}'");
                    if (ws.Cell("A2").GetValue<string>() != "3. Sor")
                        throw new Exception($"A 2. sor tartalma hibás. Várt: '3. Sor', Kapott: '{ws.Cell("A2").Value}'");
                    if (ws.Cell("A3").GetValue<string>() != "1. Sor")
                        throw new Exception($"A 3. sor tartalma hibás. Várt: '1. Sor', Kapott: '{ws.Cell("A3").Value}'");
                    if (ws.Row(3).Style.Fill.BackgroundColor.Color != XLColor.Yellow.Color)
                        throw new Exception("A stílus nem lett áthelyezve.");
                }
            });

            runTest("MoveRow: Sort felfelé mozgat (3 -> 1)", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "1. Sor";
                    ws.Cell("A2").Value = "2. Sor";
                    ws.Cell("A3").Value = "3. Sor";
                    ws.Row(3).Style.Font.Bold = true;

                    // Act
                    WorksheetExtensions.MoveRow(ws, 3, 1);

                    // Assert
                    if (ws.Cell("A1").GetValue<string>() != "3. Sor")
                        throw new Exception($"Az 1. sor tartalma hibás. Várt: '3. Sor', Kapott: '{ws.Cell("A1").Value}'");
                    if (ws.Cell("A2").GetValue<string>() != "1. Sor")
                        throw new Exception($"A 2. sor tartalma hibás. Várt: '1. Sor', Kapott: '{ws.Cell("A2").Value}'");
                    if (ws.Cell("A3").GetValue<string>() != "2. Sor")
                        throw new Exception($"A 3. sor tartalma hibás. Várt: '2. Sor', Kapott: '{ws.Cell("A3").Value}'");
                    if (!ws.Row(1).Style.Font.Bold)
                        throw new Exception("A stílus nem lett áthelyezve.");
                }
            });

            runTest("MoveRow: Érvénytelen paraméterekkel nem csinál semmit", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "A"; ws.Cell("A2").Value = "B";
                    
                    Action<int, int> checkNoChange = (source, dest) =>
                    {
                        WorksheetExtensions.MoveRow(ws, source, dest);
                        if (ws.Cell("A1").GetValue<string>() != "A" || ws.Cell("A2").GetValue<string>() != "B")
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

            runTest("MoveRow: Üres munkalapon nem okoz hibát", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");

                    // Act & Assert (nem dobhat kivételt)
                    WorksheetExtensions.MoveRow(ws, 1, 2);
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