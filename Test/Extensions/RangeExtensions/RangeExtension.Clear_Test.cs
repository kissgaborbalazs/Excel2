using System;
using System.Collections.Generic;
using System.IO;
using UiPath.CodedWorkflows;
using ClosedXML.Excel;
using System.Diagnostics;
using System.Linq;

// A tesztelendő osztály névtere.
using MPRT.Excel;

namespace MPRT.Excel.Tests
{
    public class RangeExtensions_Clear_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "RangeExtensionsClearTests", Guid.NewGuid().ToString());
            Directory.CreateDirectory(testDir);
            Log($"Temporary test directory created at: {testDir}");

            // Action a tesztek futtatásához és eredményének naplózásához
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

            // Helper metódus a tesztkörnyezet előkészítéséhez
            Action<string, Action<IXLWorksheet>> setupAndRun = (fileName, testLogic) => 
            {
                string filePath = Path.Combine(testDir, fileName);
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("TestDataSheet");
                    testLogic(worksheet);
                    workbook.SaveAs(filePath);
                }
            };
            
            // =================================================================
            // TESZTESETEK DEFINÍCIÓJA
            // =================================================================

            runTest("Clear(false): Csak a tartalmat törli, a stílust meghagyja", () =>
            {
                // Arrange
                string filePath = Path.Combine(testDir, "clear_contents_only.xlsx");
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("TestSheet");
                    var range = worksheet.Range("A1:B2");
                    range.Value = "Test Data";
                    range.Style.Font.Bold = true;
                    range.Style.Fill.BackgroundColor = XLColor.Yellow;
                    
                    // Act
                    // JAVÍTVA: Explicit hívás a bővítő metódusra.
                    RangeExtensions.Clear(range); // clearStyles = false (alapértelmezett)
                    
                    // Assert
                    foreach (var cell in range.Cells())
                    {
                        if (!cell.IsEmpty())
                            throw new Exception($"A cella tartalma nem lett törölve: {cell.Address}");
                        if (!cell.Style.Font.Bold)
                            throw new Exception($"A cella 'Bold' stílusa elveszett: {cell.Address}");
                        if (cell.Style.Fill.BackgroundColor != XLColor.Yellow)
                            throw new Exception($"A cella háttérszíne elveszett: {cell.Address}");
                    }
                    workbook.SaveAs(filePath); // Elmentjük, hogy vizuálisan is ellenőrizhető legyen
                }
            });

            runTest("Clear(true): A tartalmat és a stílust is törli", () =>
            {
                // Arrange
                string filePath = Path.Combine(testDir, "clear_all.xlsx");
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("TestSheet");
                    var range = worksheet.Range("C3:D4");
                    range.Value = "To Be Deleted";
                    range.Style.Font.Italic = true;
                    range.Style.Fill.BackgroundColor = XLColor.LightBlue;

                    // Act
                    // JAVÍTVA: Explicit hívás a bővítő metódusra.
                    RangeExtensions.Clear(range, true); // clearStyles = true

                    // Assert
                    foreach (var cell in range.Cells())
                    {
                        if (!cell.IsEmpty())
                            throw new Exception($"A cella tartalma nem lett törölve: {cell.Address}");
                        if (cell.Style.Font.Italic)
                            throw new Exception($"A cella 'Italic' stílusa nem lett törölve: {cell.Address}");
                        if (cell.Style.Fill.BackgroundColor.Color != XLColor.Transparent.Color)
                            throw new Exception($"A cella háttérszíne nem lett törölve: {cell.Address} Aktuális szín: {cell.Style.Fill.BackgroundColor.Color}");
                    }
                    workbook.SaveAs(filePath);
                }
            });

            runTest("Clear: Üres tartományon nem okoz hibát", () =>
            {
                // Arrange
                string filePath = Path.Combine(testDir, "clear_empty_range.xlsx");
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("TestSheet");
                    var range = worksheet.Range("E1:E5");

                    // Act & Assert (nem dobhat kivételt)
                    // JAVÍTVA: Explicit hívás a bővítő metódusra.
                    RangeExtensions.Clear(range);
                    RangeExtensions.Clear(range, true);
                    
                    workbook.SaveAs(filePath);
                }
            });

            runTest("Clear: Egyetlen cellán is helyesen működik", () =>
            {
                // Arrange
                string filePath = Path.Combine(testDir, "clear_single_cell.xlsx");
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("TestSheet");
                    var cellWithContent = worksheet.Cell("A1");
                    var cellWithStyle = worksheet.Cell("B2");

                    cellWithContent.Value = "Single Cell";
                    cellWithContent.Style.Font.Bold = true;

                    cellWithStyle.Value = "Delete All";
                    cellWithStyle.Style.Fill.BackgroundColor = XLColor.Red;

                    // Act
                    // JAVÍTVA: Explicit hívás a bővítő metódusra.
                    RangeExtensions.Clear(cellWithContent.AsRange()); // Csak tartalom törlése
                    RangeExtensions.Clear(cellWithStyle.AsRange(), true); // Minden törlése

                    // Assert
                    if (!cellWithContent.IsEmpty())
                        throw new Exception("Az A1 cella tartalma nem lett törölve.");
                    if (!cellWithContent.Style.Font.Bold)
                        throw new Exception("Az A1 cella stílusa hibásan törlődött.");
                    
                    if (!cellWithStyle.IsEmpty())
                        throw new Exception("A B2 cella tartalma nem lett törölve.");
                    if (cellWithStyle.Style.Fill.BackgroundColor.Color != XLColor.Transparent.Color)
                        throw new Exception("A B2 cella stílusa nem lett törölve.");

                    workbook.SaveAs(filePath);
                }
            });

            runTest("Clear: Képleteket is töröl", () =>
            {
                // Arrange
                string filePath = Path.Combine(testDir, "clear_formulas.xlsx");
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("TestSheet");
                    worksheet.Cell("A1").Value = 10;
                    worksheet.Cell("A2").Value = 20;
                    var formulaCell = worksheet.Cell("A3");
                    formulaCell.FormulaA1 = "=SUM(A1:A2)";

                    // Act
                    // JAVÍTVA: Explicit hívás a bővítő metódusra.
                    RangeExtensions.Clear(formulaCell.AsRange());

                    // Assert
                    if (!formulaCell.IsEmpty())
                        throw new Exception("A képletet tartalmazó cella tartalma nem lett törölve.");
                    if (formulaCell.HasFormula)
                        throw new Exception("A képlet nem lett törölve a cellából.");

                    workbook.SaveAs(filePath);
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