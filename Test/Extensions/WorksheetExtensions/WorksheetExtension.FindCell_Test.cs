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
    public class WorksheetExtensions_FindCell_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorksheetExtFindCellTests", Guid.NewGuid().ToString());
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

            runTest("FindCell: Megtalál egy létező string értéket", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("C5").Value = "Keresendő Érték";

                    // Act
                    var address = WorksheetExtensions.FindCell(ws, "Keresendő Érték");

                    // Assert
                    if (address != "C5")
                        throw new Exception($"Hibás cellacímet adott vissza. Várt: 'C5', Kapott: '{address}'");
                }
            });

            runTest("FindCell: Üres stringet ad vissza nem létező értékre", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Valami más";

                    // Act
                    var address = WorksheetExtensions.FindCell(ws, "Nem Létező Érték");

                    // Assert
                    if (address != string.Empty)
                        throw new Exception($"Nem üres stringet adott vissza. Várt: '', Kapott: '{address}'");
                }
            });

            runTest("FindCell: Kis- és nagybetűre érzéketlenül keres", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("B2").Value = "Case Test";

                    // Act
                    var address = WorksheetExtensions.FindCell(ws, "case test");

                    // Assert
                    if (address != "B2")
                        throw new Exception($"A keresés nem volt érzéketlen a kis- és nagybetűkre. Várt: 'B2', Kapott: '{address}'");
                }
            });

            runTest("FindCell: Megtalál egy számot is", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("D10").Value = 123.45;

                    // Act
                    var address = WorksheetExtensions.FindCell(ws, 123.45);

                    // Assert
                    if (address != "D10")
                        throw new Exception($"Nem találta meg a számot. Várt: 'D10', Kapott: '{address}'");
                }
            });
            
            runTest("FindCell: Az első előfordulást adja vissza", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Ismétlődő";
                    ws.Cell("F8").Value = "Ismétlődő";

                    // Act
                    var address = WorksheetExtensions.FindCell(ws, "Ismétlődő");

                    // Assert
                    if (address != "A1")
                        throw new Exception($"Nem az első előfordulást adta vissza. Várt: 'A1', Kapott: '{address}'");
                }
            });

            runTest("FindCell: Üres stringet ad vissza null inputra", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Valami";

                    // Act
                    var address = WorksheetExtensions.FindCell(ws, null);

                    // Assert
                    if (address != string.Empty)
                        throw new Exception($"Nem üres stringet adott vissza null inputra. Várt: '', Kapott: '{address}'");
                }
            });

            runTest("FindCell: Üres munkalapon üres stringet ad vissza", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");

                    // Act
                    var address = WorksheetExtensions.FindCell(ws, "Bármi");

                    // Assert
                    if (address != string.Empty)
                        throw new Exception($"Nem üres stringet adott vissza üres munkalapon. Várt: '', Kapott: '{address}'");
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