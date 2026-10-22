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
    public class WorkbookExtensions_SheetExists_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorkbookExtSheetExistsTests", Guid.NewGuid().ToString());
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

            runTest("SheetExists: Igazat ad vissza létező munkalapra", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    workbook.Worksheets.Add("Létező Lap");
                    
                    // Act
                    var exists = WorkbookExtensions.SheetExists(workbook, "Létező Lap");

                    // Assert
                    if (!exists)
                    {
                        throw new Exception("A metódus hamisat adott vissza egy létező lapra. Várt: true, Kapott: false");
                    }
                }
            });

            runTest("SheetExists: Hamisat ad vissza nem létező munkalapra", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    workbook.Worksheets.Add("Valami Más");

                    // Act
                    var exists = WorkbookExtensions.SheetExists(workbook, "Nem Létező Lap");

                    // Assert
                    if (exists)
                    {
                        throw new Exception("A metódus igazat adott vissza egy nem létező lapra. Várt: false, Kapott: true");
                    }
                }
            });

            runTest("SheetExists: Kis- és nagybetűre érzéketlen", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    workbook.Worksheets.Add("MySheet");

                    // Act
                    var existsLowerCase = WorkbookExtensions.SheetExists(workbook, "mysheet");
                    var existsUpperCase = WorkbookExtensions.SheetExists(workbook, "MYSHEET");

                    // Assert
                    if (!existsLowerCase)
                    {
                        throw new Exception("A keresés nem volt érzéketlen a kisbetűkre. Várt: true, Kapott: false");
                    }
                    if (!existsUpperCase)
                    {
                        throw new Exception("A keresés nem volt érzéketlen a nagybetűkre. Várt: true, Kapott: false");
                    }
                }
            });

            runTest("SheetExists: Hamisat ad vissza null, üres vagy whitespace stringre", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    workbook.Worksheets.Add("Valódi Lap");

                    // Act
                    var existsNull = WorkbookExtensions.SheetExists(workbook, null);
                    var existsEmpty = WorkbookExtensions.SheetExists(workbook, "");
                    var existsWhitespace = WorkbookExtensions.SheetExists(workbook, "   ");

                    // Assert
                    if (existsNull)
                    {
                        throw new Exception("A metódus igazat adott vissza null inputra. Várt: false, Kapott: true");
                    }
                    if (existsEmpty)
                    {
                        throw new Exception("A metódus igazat adott vissza üres string inputra. Várt: false, Kapott: true");
                    }
                    if (existsWhitespace)
                    {
                        throw new Exception("A metódus igazat adott vissza csak whitespace-t tartalmazó inputra. Várt: false, Kapott: true");
                    }
                }
            });

            runTest("SheetExists: Hamisat ad vissza üres munkafüzetben", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    // Act
                    var exists = WorkbookExtensions.SheetExists(workbook, "Bármi");

                    // Assert
                    if (exists)
                    {
                        throw new Exception("A metódus igazat adott vissza egy üres munkafüzetben. Várt: false, Kapott: true");
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