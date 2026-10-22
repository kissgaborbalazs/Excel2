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
    public class WorkbookExtensions_GetSheetNames_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorkbookExtGetSheetNamesTests", Guid.NewGuid().ToString());
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

            runTest("GetSheetNames: Új, üres munkafüzet esetén üres listát ad vissza", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    // Act
                    var sheetNames = WorkbookExtensions.GetSheetNames(workbook);

                    // Assert
                    if (sheetNames.Any())
                    {
                        throw new Exception($"A lista nem üres. Tartalma: {string.Join(", ", sheetNames)}");
                    }
                }
            });

            runTest("GetSheetNames: Egy munkalap esetén a helyes nevet adja vissza", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    workbook.Worksheets.Add("Első Lap");
                    
                    // Act
                    var sheetNames = WorkbookExtensions.GetSheetNames(workbook).ToList();

                    // Assert
                    if (sheetNames.Count != 1 || sheetNames[0] != "Első Lap")
                    {
                        throw new Exception($"Hibás eredménylista. Várt: ['Első Lap'], Kapott: [{string.Join(", ", sheetNames)}]");
                    }
                }
            });

            runTest("GetSheetNames: Több munkalap esetén minden nevet visszaad a helyes sorrendben", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var expectedNames = new List<string> { "Adatok", "Kimutatás", "Forrás" };
                    foreach (var name in expectedNames)
                    {
                        workbook.Worksheets.Add(name);
                    }
                    
                    // Act
                    var actualNames = WorkbookExtensions.GetSheetNames(workbook).ToList();

                    // Assert
                    if (!expectedNames.SequenceEqual(actualNames))
                    {
                        throw new Exception($"A munkalapnevek listája vagy sorrendje hibás. Várt: {string.Join(", ", expectedNames)}, Kapott: {string.Join(", ", actualNames)}");
                    }
                }
            });

            runTest("GetSheetNames: Különleges karaktereket és szóközöket tartalmazó neveket is helyesen kezel", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var expectedNames = new List<string> { "Árvíztűrő lap", "adatok-2025 (verzió 2)" };
                    workbook.Worksheets.Add(expectedNames[0]);
                    workbook.Worksheets.Add(expectedNames[1]);
                    
                    // Act
                    var actualNames = WorkbookExtensions.GetSheetNames(workbook).ToList();

                    // Assert
                    if (actualNames.Count != 2 || !actualNames.Contains(expectedNames[0]) || !actualNames.Contains(expectedNames[1]))
                    {
                        throw new Exception($"A speciális karaktereket tartalmazó lapnevek nem megfelelőek. Várt: {string.Join(", ", expectedNames)}, Kapott: {string.Join(", ", actualNames)}");
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