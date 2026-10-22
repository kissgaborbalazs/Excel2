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
    public class WorksheetExtensions_WriteDataTable_Tests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "WorksheetExtWriteDtTests", Guid.NewGuid().ToString());
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

            runTest("WriteDataTable: Helyesen beírja az adatokat fejléccel együtt", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    var dt = new DataTable();
                    dt.Columns.Add("ID", typeof(int));
                    dt.Columns.Add("Név", typeof(string));
                    dt.Rows.Add(1, "Anna");

                    // Act
                    WorksheetExtensions.WriteDataTable(ws, dt, "A1", true);

                    // Assert
                    if (ws.Cell("A1").GetValue<string>() != "ID" || ws.Cell("B1").GetValue<string>() != "Név")
                        throw new Exception("A fejlécek nem íródtak be helyesen.");
                    if (ws.Cell("A2").GetValue<int>() != 1 || ws.Cell("B2").GetValue<string>() != "Anna")
                        throw new Exception("Az adatok nem íródtak be helyesen.");
                }
            });

            runTest("WriteDataTable: Helyesen beírja az adatokat fejléc nélkül", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    var dt = new DataTable();
                    dt.Columns.Add("ID", typeof(int));
                    dt.Rows.Add(10);

                    // Act
                    WorksheetExtensions.WriteDataTable(ws, dt, "A1", false);

                    // Assert
                    if (ws.Cell("A1").GetValue<int>() != 10)
                        throw new Exception($"Az adat fejléc nélkül nem íródott be helyesen. Várt: 10, Kapott: {ws.Cell("A1").Value}");
                    if (ws.LastCellUsed().Address.ToString() != "A1")
                        throw new Exception("Több cella lett beírva a vártnál.");
                }
            });

            runTest("WriteDataTable: Adott cellától kezdi az írást", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    var dt = new DataTable();
                    dt.Columns.Add("Adat");
                    dt.Rows.Add("Teszt");

                    // Act
                    WorksheetExtensions.WriteDataTable(ws, dt, "C5", true);

                    // Assert
                    if (ws.Cell("C5").GetValue<string>() != "Adat")
                        throw new Exception("A fejléc nem a megfelelő helyre került.");
                    if (ws.Cell("C6").GetValue<string>() != "Teszt")
                        throw new Exception("Az adat nem a megfelelő helyre került.");
                }
            });

            runTest("WriteDataTable: Üres DataTable esetén nem csinál semmit", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");
                    var dt = new DataTable(); // Nincsenek sorok

                    // Act
                    WorksheetExtensions.WriteDataTable(ws, dt);

                    // Assert
                    if (ws.LastCellUsed() != null)
                        throw new Exception("A munkalap nem maradt üres.");
                }
            });

            runTest("WriteDataTable: Null DataTable esetén nem csinál semmit", () =>
            {
                // Arrange
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("TestSheet");

                    // Act
                    WorksheetExtensions.WriteDataTable(ws, null);

                    // Assert
                    if (ws.LastCellUsed() != null)
                        throw new Exception("A munkalap nem maradt üres.");
                }
            });

            runTest("WriteDataTable: Null worksheet esetén ArgumentNullException-t dob", () =>
            {
                // Arrange
                IXLWorksheet nullSheet = null;
                var dt = new DataTable();

                // Act & Assert
                try
                {
                    WorksheetExtensions.WriteDataTable(nullSheet, dt);
                    throw new Exception("Nem dobódott a várt ArgumentNullException.");
                }
                catch (ArgumentNullException)
                {
                    /* Test Passed */
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
