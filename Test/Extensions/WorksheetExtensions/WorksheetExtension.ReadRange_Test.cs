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
    public class WorksheetExtensions_ReadRange_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorksheetExtReadRangeTests", Guid.NewGuid().ToString());
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

            runTest("ReadRange: Teljes használt területet olvas (rangeAddress=null)", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("B2").Value = "ID"; ws.Cell("C2").Value = "Név";
                    ws.Cell("B3").Value = 1; ws.Cell("C3").Value = "Anna";
                    
                    // Act
                    var dt = WorksheetExtensions.ReadRange(ws, null, true);

                    // Assert
                    if (dt.Columns.Count != 2 || dt.Columns[0].ColumnName != "ID" || dt.Columns[1].ColumnName != "Név")
                        throw new Exception("A fejlécek beolvasása hibás.");
                    if (dt.Rows.Count != 1 || dt.Rows[0]["ID"].ToString() != "1")
                        throw new Exception("Az adatsorok beolvasása hibás.");
                }
            });

            runTest("ReadRange: Explicit tartományt olvas fejléc nélkül", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Fejléc";
                    ws.Cell("A2").Value = "Adat1";
                    ws.Cell("A3").Value = "Adat2";
                    
                    // Act
                    var dt = WorksheetExtensions.ReadRange(ws, "A2:A3", false);

                    // Assert
                    if (dt.Columns[0].ColumnName != "Column1")
                        throw new Exception("A generált oszlopnév hibás.");
                    if (dt.Rows.Count != 2 || dt.Rows[0][0].ToString() != "Adat1")
                        throw new Exception("Az adatok beolvasása fejléc nélkül hibás.");
                }
            });

            runTest("ReadRange: Kezdőcellából automatikusan felismeri a tartományt", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("B2").Value = "Fejléc1"; ws.Cell("C2").Value = "Fejléc2";
                    ws.Cell("B3").Value = "a"; ws.Cell("C3").Value = "b";
                    ws.Cell("B4").Value = "c"; ws.Cell("C4").Value = "d";
                    
                    // Act
                    var dt = WorksheetExtensions.ReadRange(ws, "B2", true);

                    // Assert
                    if (dt.Columns.Count != 2 || dt.Rows.Count != 2)
                        throw new Exception($"Az automatikus tartományfelismerés hibás. Oszlopok: {dt.Columns.Count}, Sorok: {dt.Rows.Count}");
                    if (dt.Rows[1]["Fejléc2"].ToString() != "d")
                        throw new Exception("Az utolsó cella értéke hibás.");
                }
            });

            runTest("ReadRange: keepFormatting=true a formázott értéket olvassa", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Dátum";
                    ws.Cell("A2").Value = new DateTime(2025, 8, 15);
                    ws.Cell("A2").Style.DateFormat.Format = "yyyy-MM-dd";
                    
                    // Act
                    var dt = WorksheetExtensions.ReadRange(ws, "A1:A2", true, true);

                    // Assert
                    var formattedValue = dt.Rows[0][0].ToString();
                    if (formattedValue != "2025-08-15")
                        throw new Exception($"A formázott érték beolvasása hibás. Várt: '2025-08-15', Kapott: '{formattedValue}'");
                }
            });

            runTest("ReadRange: keepFormatting=false a nyers értéket olvassa", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    ws.Cell("A1").Value = "Dátum";
                    var testDate = new DateTime(2025, 8, 15);
                    ws.Cell("A2").Value = testDate;
                    ws.Cell("A2").Style.DateFormat.Format = "yyyy-MM-dd";
                    
                    // Act
                    var dt = WorksheetExtensions.ReadRange(ws, "A1:A2", true, false);

                    // Assert
                    if (dt.Rows[0][0].GetType() != typeof(DateTime))
                        throw new Exception($"A nyers érték típusa hibás. Várt: DateTime, Kapott: {dt.Rows[0][0].GetType()}");
                    if ((DateTime)dt.Rows[0][0] != testDate)
                        throw new Exception("A nyers dátumérték hibás.");
                }
            });
            
            runTest("ReadRange: Üres munkalapról üres DataTable-t ad vissza", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    
                    // Act
                    var dt = WorksheetExtensions.ReadRange(ws);

                    // Assert
                    if (dt.Rows.Count != 0 || dt.Columns.Count != 0)
                        throw new Exception("Nem üres DataTable-t adott vissza üres munkalapra.");
                }
            });

            runTest("ReadRange: Null worksheet esetén ArgumentNullException-t dob", () =>
            {
                // Arrange
                IXLWorksheet nullSheet = null;

                // Act & Assert
                try
                {
                    WorksheetExtensions.ReadRange(nullSheet);
                    throw new Exception("Nem dobódott a várt ArgumentNullException.");
                }
                catch (ArgumentNullException) { /* Test Passed */ }
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
