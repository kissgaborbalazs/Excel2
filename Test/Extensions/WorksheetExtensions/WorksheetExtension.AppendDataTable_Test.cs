using System;
using System.IO;
using UiPath.CodedWorkflows;
using ClosedXML.Excel;
using System.Diagnostics;
using System.Linq;
using System.Collections.Generic;
using System.Data;

// A tesztelendő osztály névtere.
using MPRT.Excel;

namespace MPRT.Excel.Tests
{
    public class WorksheetExtensions_AppendDataTable_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorksheetExtAppendDtTests", Guid.NewGuid().ToString());
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

            runTest("AppendDataTable: Hozzáfűzi az adatokat a meglévő tartalom után", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "ID"; ws.Cell("B1").Value = "Név";
                    ws.Cell("A2").Value = 1; ws.Cell("B2").Value = "Első";
                    
                    var dt = new DataTable();
                    dt.Columns.Add("ID", typeof(int));
                    dt.Columns.Add("Név", typeof(string));
                    dt.Rows.Add(2, "Második");
                    dt.Rows.Add(3, "Harmadik");

                    // Act
                    WorksheetExtensions.AppendDataTable(ws, dt);

                    // Assert
                    if (ws.Cell("A3").GetValue<int>() != 2 || ws.Cell("B3").GetValue<string>() != "Második")
                        throw new Exception($"A hozzáfűzés első sora hibás. Várt: [2, Második], Kapott: [{ws.Cell("A3").Value}, {ws.Cell("B3").Value}]");
                    if (ws.Cell("A4").GetValue<int>() != 3 || ws.Cell("B4").GetValue<string>() != "Harmadik")
                        throw new Exception($"A hozzáfűzés második sora hibás. Várt: [3, Harmadik], Kapott: [{ws.Cell("A4").Value}, {ws.Cell("B4").Value}]");
                    if (ws.LastCellUsed().Address.RowNumber != 4)
                        throw new Exception($"Az utolsó használt sor hibás. Várt: 4, Kapott: {ws.LastCellUsed().Address.RowNumber}");
                }
            });

            runTest("AppendDataTable: Üres munkalaphoz is helyesen fűz hozzá", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    var dt = new DataTable();
                    dt.Columns.Add("TesztOszlop");
                    dt.Rows.Add("TesztAdat");

                    // Act
                    // A WriteDataTable stub miatt A2-be fog írni, mert a startCell A1 lesz.
                    WorksheetExtensions.AppendDataTable(ws, dt);

                    // Assert
                    if (ws.Cell("A1").GetValue<string>() != "TesztAdat")
                        throw new Exception($"Az adat nem a várt helyre került. Várt 'TesztAdat' az A2 cellában, de a kapott: '{ws.Cell("A2").Value}'");
                }
            });

            runTest("AppendDataTable: Üres DataTable esetén nem változtat a munkalapon", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Eredeti adat";
                    var lastRowBefore = ws.LastCellUsed()?.Address.RowNumber ?? 0;
                    
                    var dt = new DataTable(); // Üres, nincsenek sorok

                    // Act
                    WorksheetExtensions.AppendDataTable(ws, dt);

                    // Assert
                    var lastRowAfter = ws.LastCellUsed()?.Address.RowNumber ?? 0;
                    if (lastRowBefore != lastRowAfter)
                        throw new Exception($"A munkalap hibásan megváltozott. Sorok előtte: {lastRowBefore}, Sorok utána: {lastRowAfter}");
                }
            });

            runTest("AppendDataTable: Null DataTable esetén nem változtat a munkalapon", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Eredeti adat";
                    var lastRowBefore = ws.LastCellUsed()?.Address.RowNumber ?? 0;

                    // Act
                    WorksheetExtensions.AppendDataTable(ws, null);

                    // Assert
                    var lastRowAfter = ws.LastCellUsed()?.Address.RowNumber ?? 0;
                    if (lastRowBefore != lastRowAfter)
                        throw new Exception($"A munkalap hibásan megváltozott. Sorok előtte: {lastRowBefore}, Sorok utána: {lastRowAfter}");
                }
            });

            runTest("AppendDataTable: Különböző adattípusokat is helyesen kezel", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Szöveg"; ws.Cell("B1").Value = "Szám"; ws.Cell("C1").Value = "Dátum";
                    
                    var dt = new DataTable();
                    dt.Columns.Add("ColA", typeof(string));
                    dt.Columns.Add("ColB", typeof(int));
                    dt.Columns.Add("ColC", typeof(DateTime));
                    var testDate = new DateTime(2025, 8, 15);
                    dt.Rows.Add("Teszt", 123, testDate);

                    // Act
                    WorksheetExtensions.AppendDataTable(ws, dt);

                    // Assert
                    if (ws.Cell("A2").GetValue<string>() != "Teszt") throw new Exception("A string adat hibás.");
                    if (ws.Cell("B2").GetValue<int>() != 123) throw new Exception("Az integer adat hibás.");
                    if (ws.Cell("C2").GetValue<DateTime>() != testDate) throw new Exception("A DateTime adat hibás.");
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