using System;
using System.Collections.Generic;
using System.IO;
using UiPath.CodedWorkflows;
using ClosedXML.Excel;
using System.Diagnostics;
using System.Linq;

// A tesztelendő osztály névtere.
// Fontos, hogy a tesztek hozzáférjenek a belső (internal) elemekhez,
// ha szükséges. Ezt a tesztelendő projekt .csproj fájljában az
// <InternalsVisibleTo Include="Your.Test.Project.Name" /> beállítással lehet engedélyezni.
using MPRT.Excel;

namespace MPRT.Excel.Tests
{
    public class ExcelManager_CodedTests : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
            var passed = 0;
            var failed = 0;
            var total = 0;
            
            string testDir = Path.Combine(Path.GetTempPath(), "ExcelManagerTests", Guid.NewGuid().ToString());
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
            
            // =================================================================
            // TESZTESETEK DEFINÍCIÓJA
            // =================================================================

            #region Constructor Tests

            runTest("Konstruktor: Új fájl létrehozása, ha nem létezik", () =>
            {
                string filePath = Path.Combine(testDir, "new_file.xlsx");
                using (var excel = new ExcelManager(filePath))
                {
                    excel.Worksheet("DefaultSheet");
                }

                // Assert
                if (!File.Exists(filePath))
                    throw new Exception("A fájl nem jött létre a Dispose metódus hívása után.");
            });
            
            runTest("Konstruktor: Létező fájl sikeres megnyitása", () =>
            {
                // Arrange
                string filePath = Path.Combine(testDir, "existing_file.xlsx");
                using (var wb = new XLWorkbook())
                {
                    wb.AddWorksheet("TestSheet");
                    wb.SaveAs(filePath);
                }

                // Act & Assert (nem dobhat kivételt)
                using (var excel = new ExcelManager(filePath))
                {
                   var sheet = excel.Worksheet("TestSheet");
                   if(sheet == null)
                       throw new Exception("A létező munkalapot nem sikerült betölteni.");
                }
            });

            runTest("Konstruktor: ArgumentNullException dobása null elérési útvonalra", () =>
            {
                try
                {
                    new ExcelManager(null);
                    throw new Exception("Nem dobódott a várt ArgumentNullException.");
                }
                catch (ArgumentNullException) { /* Test Passed */ }
            });

            runTest("Konstruktor: ArgumentNullException dobása üres string elérési útvonalra", () =>
            {
                try
                {
                    new ExcelManager("");
                    throw new Exception("Nem dobódott a várt ArgumentNullException.");
                }
                catch (ArgumentNullException) { /* Test Passed */ }
            });
            
            runTest("Konstruktor: ArgumentNullException dobása csak whitespace-t tartalmazó útvonalra", () =>
            {
                try
                {
                    new ExcelManager("   ");
                    throw new Exception("Nem dobódott a várt ArgumentNullException.");
                }
                catch (ArgumentNullException) { /* Test Passed */ }
            });

            runTest("Konstruktor: DirectoryNotFoundException dobása nem létező mappára", () =>
            {
                // Arrange
                string invalidPath = Path.Combine(testDir, "non_existent_dir", "file.xlsx");

                // Act & Assert
                try
                {
                    new ExcelManager(invalidPath);
                    throw new Exception("Nem dobódott a várt DirectoryNotFoundException.");
                }
                catch (DirectoryNotFoundException) { /* Test Passed */ }
            });

            runTest("Konstruktor: ArgumentException dobása, ha az útvonal egy mappa", () =>
            {
                // Arrange
                string dirPath = Path.Combine(testDir, "a_directory");
                Directory.CreateDirectory(dirPath);

                // Act & Assert
                try
                {
                    new ExcelManager(dirPath);
                    throw new Exception("Nem dobódott a várt ArgumentException.");
                }
                catch (ArgumentException) { /* Test Passed */ }
            });

            runTest("Konstruktor: Hibás (nem Excel) fájl megnyitásakor kivételt dob", () =>
            {
                // Arrange
                string filePath = Path.Combine(testDir, "not_an_excel.xlsx");
                File.WriteAllText(filePath, "Ez egy sima text file, nem pedig egy ZIP alapú XLSX.");

                // Act & Assert
                try
                {
                    new ExcelManager(filePath);
                    throw new Exception("A ClosedXML nem dobott kivételt a hibás fájlformátumra.");
                }
                catch (Exception ex)
                {
                    if (ex is FileFormatException)
                    {
                        /* Test Passed */
                        Log("Helyes kivétel elkapva: IOException (Inner: InvalidDataException).");
                    }
                    else
                    {
                        throw new Exception($"Váratlan kivételtípus: '{ex.GetType().FullName}'. A várt egy formátummal kapcsolatos hiba lett volna (pl. IOException).", ex);
                    }
                }
            });
            
            #endregion

            #region Worksheet Method Tests

            runTest("Worksheet metódus: Visszaadja a létező munkalapot", () =>
            {
                // Arrange
                string filePath = Path.Combine(testDir, "sheet_exists.xlsx");
                using (var wb = new XLWorkbook())
                {
                    wb.AddWorksheet("MySheet").Cell("A1").Value = "Marker";
                    wb.SaveAs(filePath);
                }

                // Act
                using (var excel = new ExcelManager(filePath))
                {
                    var sheet = excel.Worksheet("MySheet");
                    // Assert
                    if (sheet.Cell("A1").GetValue<string>() != "Marker")
                        throw new Exception("Nem a megfelelő munkalapot vagy nem a helyes tartalommal adta vissza.");
                }
            });

            runTest("Worksheet metódus: Létrehozza a munkalapot, ha nem létezik", () =>
            {
                // Arrange
                string filePath = Path.Combine(testDir, "create_sheet.xlsx");
                
                // Act
                using (var excel = new ExcelManager(filePath))
                {
                    excel.Worksheet("NewlyCreatedSheet").Cell("A1").Value = "Success";
                }

                // Assert
                using (var resultWb = new XLWorkbook(filePath))
                {
                    if (!resultWb.Worksheets.TryGetWorksheet("NewlyCreatedSheet", out var sheet))
                        throw new Exception("A munkalap nem jött létre a mentés során.");
                    
                    if(sheet.Cell("A1").GetValue<string>() != "Success")
                        throw new Exception("A létrehozott munkalapra írt adat helytelen.");
                }
            });

            runTest("Worksheet metódus: ArgumentNullException dobása null munkalapnévre", () =>
            {
                string filePath = Path.Combine(testDir, "null_sheet_name.xlsx");
                using (var excel = new ExcelManager(filePath))
                {
                    try
                    {
                        excel.Worksheet(null);
                        throw new Exception("Nem dobódott a várt ArgumentNullException.");
                    }
                    catch (ArgumentNullException)
                    {
                        excel.Worksheet("dummy_sheet_for_dispose");
                    }
                }
            });

            runTest("Worksheet metódus: ArgumentNullException dobása üres munkalapnévre", () =>
            {
                string filePath = Path.Combine(testDir, "empty_sheet_name.xlsx");
                using (var excel = new ExcelManager(filePath))
                {
                    try
                    {
                        excel.Worksheet("   ");
                        throw new Exception("Nem dobódott a várt ArgumentNullException.");
                    }
                    catch (ArgumentNullException)
                    {
                        excel.Worksheet("dummy_sheet_for_dispose");
                    }
                }
            });
            
            #endregion

            #region GetSheetNames Method Tests

            runTest("GetSheetNames: Új, üres munkafüzet esetén üres listát ad vissza", () =>
            {
                string filePath = Path.Combine(testDir, "getsheets_empty.xlsx");
                using(var excel = new ExcelManager(filePath))
                {
                    var sheetNames = excel.GetSheetNames();
                    if (sheetNames.Any())
                    {
                        throw new Exception("A GetSheetNames nem üres listát adott vissza egy üres munkafüzetre.");
                    }
                    // A Dispose miatt hozzá kell adni egy lapot, hogy menthető legyen.
                    excel.Worksheet("dummy");
                }
            });

            runTest("GetSheetNames: Egy munkalap esetén a helyes nevet adja vissza", () =>
            {
                string filePath = Path.Combine(testDir, "getsheets_one.xlsx");
                using(var excel = new ExcelManager(filePath))
                {
                    excel.Worksheet("Első Lap");
                    var sheetNames = excel.GetSheetNames().ToList();
                    if (sheetNames.Count != 1 || sheetNames[0] != "Első Lap")
                    {
                        throw new Exception($"Várt egyetlen elem 'Első Lap', de a kapott: {string.Join(", ", sheetNames)}");
                    }
                }
            });

            runTest("GetSheetNames: Több munkalap esetén minden nevet visszaad", () =>
            {
                string filePath = Path.Combine(testDir, "getsheets_multiple.xlsx");
                var expectedNames = new List<string> { "Adatok", "Kimutatás", "Forrás" };
                using(var excel = new ExcelManager(filePath))
                {
                    foreach(var name in expectedNames)
                    {
                        excel.Worksheet(name);
                    }

                    var actualNames = excel.GetSheetNames().ToList();
                    
                    if (!expectedNames.SequenceEqual(actualNames))
                    {
                        throw new Exception($"A várt és kapott lapnevek nem egyeznek. Várt: {string.Join(", ", expectedNames)}, Kapott: {string.Join(", ", actualNames)}");
                    }
                }
            });

            runTest("GetSheetNames: Különleges karaktereket és szóközöket tartalmazó neveket is helyesen kezel", () =>
            {
                string filePath = Path.Combine(testDir, "getsheets_special_chars.xlsx");
                var expectedNames = new List<string> { "Árvíztűrő lap", "adatok-2025 (verzió 2)" };
                using(var excel = new ExcelManager(filePath))
                {
                    excel.Worksheet(expectedNames[0]);
                    excel.Worksheet(expectedNames[1]);

                    var actualNames = excel.GetSheetNames().ToList();
                    
                    if (actualNames.Count != 2 || !actualNames.Contains(expectedNames[0]) || !actualNames.Contains(expectedNames[1]))
                    {
                        throw new Exception($"A speciális karaktereket tartalmazó lapnevek nem megfelelőek. Várt: {string.Join(", ", expectedNames)}, Kapott: {string.Join(", ", actualNames)}");
                    }
                }
            });

            #endregion

            #region SheetExists Method Tests

            runTest("SheetExists: Igazat ad vissza létező munkalapra", () =>
            {
                string filePath = Path.Combine(testDir, "sheetexists_true.xlsx");
                using (var excel = new ExcelManager(filePath))
                {
                    excel.Worksheet("Létező Lap");
                    if (!excel.SheetExists("Létező Lap"))
                    {
                        throw new Exception("A SheetExists hamisat adott vissza egy létező lapra.");
                    }
                }
            });

            runTest("SheetExists: Hamisat ad vissza nem létező munkalapra", () =>
            {
                string filePath = Path.Combine(testDir, "sheetexists_false.xlsx");
                using (var excel = new ExcelManager(filePath))
                {
                    excel.Worksheet("Valami Más");
                    if (excel.SheetExists("Nem Létező Lap"))
                    {
                        throw new Exception("A SheetExists igazat adott vissza egy nem létező lapra.");
                    }
                }
            });

            runTest("SheetExists: Hamisat ad vissza üres munkafüzetben", () =>
            {
                string filePath = Path.Combine(testDir, "sheetexists_empty_wb.xlsx");
                using (var excel = new ExcelManager(filePath))
                {
                    if (excel.SheetExists("Bármi"))
                    {
                        throw new Exception("A SheetExists igazat adott vissza egy üres munkafüzetben.");
                    }
                    // A Dispose miatt hozzá kell adni egy lapot, hogy menthető legyen.
                    excel.Worksheet("dummy");
                }
            });

            runTest("SheetExists: Kis- és nagybetűre érzéketlen", () =>
            {
                string filePath = Path.Combine(testDir, "sheetexists_case.xlsx");
                using (var excel = new ExcelManager(filePath))
                {
                    excel.Worksheet("MySheet");
                    if (!excel.SheetExists("mysheet"))
                    {
                        throw new Exception("A SheetExists nem volt érzéketlen a kis- és nagybetűkre (le kellett volna találnia).");
                    }
                    if (!excel.SheetExists("MYSHEET"))
                    {
                        throw new Exception("A SheetExists nem volt érzéketlen a kis- és nagybetűkre (le kellett volna találnia).");
                    }
                }
            });

            runTest("SheetExists: Hamisat ad vissza null vagy üres stringre", () =>
            {
                string filePath = Path.Combine(testDir, "sheetexists_null_empty.xlsx");
                using (var excel = new ExcelManager(filePath))
                {
                    excel.Worksheet("Valódi Lap");
                    if (excel.SheetExists(null))
                    {
                        throw new Exception("A SheetExists igazat adott vissza null inputra.");
                    }
                    if (excel.SheetExists(""))
                    {
                        throw new Exception("A SheetExists igazat adott vissza üres string inputra.");
                    }
                    if (excel.SheetExists("   "))
                    {
                        throw new Exception("A SheetExists igazat adott vissza csak whitespace-t tartalmazó inputra.");
                    }
                }
            });

            #endregion

            #region Dispose / Save Tests

            runTest("Dispose: Elmenti az új fájlba írt adatokat", () =>
            {
                // Arrange
                string filePath = Path.Combine(testDir, "save_on_dispose.xlsx");
                
                // Act
                using (var excel = new ExcelManager(filePath))
                {
                    excel.Worksheet("DataSheet").Cell("B2").Value = "Test Data";
                }

                // Assert
                using (var resultWb = new XLWorkbook(filePath))
                {
                    var val = resultWb.Worksheet("DataSheet").Cell("B2").GetValue<string>();
                    if (val != "Test Data")
                        throw new Exception($"Hibás adat a mentett fájlban. Várt: 'Test Data', Kapott: '{val}'");
                }
            });

            runTest("Dispose: Elmenti a módosításokat egy létező fájlban", () =>
            {
                // Arrange
                string filePath = Path.Combine(testDir, "modify_and_save.xlsx");
                using (var wb = new XLWorkbook())
                {
                    wb.AddWorksheet("Original").Cell("A1").Value = "Initial";
                    wb.SaveAs(filePath);
                }

                // Act
                using (var excel = new ExcelManager(filePath))
                {
                    excel.Worksheet("Original").Cell("A1").Value = "Modified";
                    excel.Worksheet("Added").Cell("C3").Value = "New Data";
                }

                // Assert
                using (var resultWb = new XLWorkbook(filePath))
                {
                    var val1 = resultWb.Worksheet("Original").Cell("A1").GetValue<string>();
                    var val2 = resultWb.Worksheet("Added").Cell("C3").GetValue<string>();
                    if (val1 != "Modified")
                        throw new Exception("A meglévő adat módosítása nem került mentésre.");
                    if (val2 != "New Data")
                        throw new Exception("Az új munkalapra írt adat nem került mentésre.");
                }
            });
            
            runTest("Dispose: Képleteket kiértékel mentés előtt", () =>
            {
                // Arrange
                string filePath = Path.Combine(testDir, "formula_evaluation.xlsx");
                
                // Act
                using (var excel = new ExcelManager(filePath))
                {
                    var sheet = excel.Worksheet("Calculations");
                    sheet.Cell("A1").Value = 10;
                    sheet.Cell("A2").Value = 20;
                    sheet.Cell("A3").FormulaA1 = "=SUM(A1:A2)";
                }

                // Assert
                using (var resultWb = new XLWorkbook(filePath))
                {
                    var sheet = resultWb.Worksheet("Calculations");
                    var resultValue = sheet.Cell("A3").GetDouble();
                    if (resultValue != 30)
                    {
                        throw new Exception($"A képlet kiértékelése helytelen. Várt: 30, Kapott: {resultValue}");
                    }
                }
            });

            runTest("Dispose: Kivételt dob írásvédett fájlba mentéskor", () =>
            {
                // Arrange
                string filePath = Path.Combine(testDir, "readonly_dispose.xlsx");
                // 1. Létrehozunk egy valid, nem üres Excel fájlt.
                using (var wb = new XLWorkbook())
                {
                    wb.AddWorksheet("sheet1").Cell("A1").Value = "Initial Data";
                    wb.SaveAs(filePath);
                }
                // 2. Írásvédetté tesszük.
                File.SetAttributes(filePath, FileAttributes.ReadOnly);

                try
                {
                    // 3. Act:
                    // A konstruktornak sikeresnek kell lennie.
                    using (var excel = new ExcelManager(filePath))
                    {
                        // Módosítunk valamit a memóriában.
                        excel.Worksheet("sheet1").Cell("A1").Value = "Modified Data";
                    } // A hiba a 'using' blokk végén, a Dispose() hívásakor kell, hogy történjen.

                    // 4. Assert: Ha ide eljut a kód, a teszt megbukott.
                    throw new Exception("A Dispose metódus nem dobott kivételt írásvédett fájlba mentéskor.");
                }
                catch (UnauthorizedAccessException)
                {
                    // A várt kivétel megérkezett a mentési kísérletkor. A teszt sikeres.
                    /* Test Passed */
                    Log("Helyes UnauthorizedAccessException elkapva a Dispose metódusban írásvédett fájl esetén.");
                }
                catch (IOException ex)
                {
                    // Néhány rendszeren vagy .NET verzióban az UnauthorizedAccessException IOException-ként
                    // csomagolva jelenhet meg. Ez a catch blokk robusztusabbá teszi a tesztet.
                    if (ex.Message.Contains("denied"))
                    {
                        /* Test Passed */
                        Log("Helyes IOException (access denied) elkapva a Dispose metódusban írásvédett fájl esetén.");
                    }
                    else
                    {
                        throw; // Ha más típusú IOException, akkor az hiba.
                    }
                }
                finally
                {
                    // Cleanup
                    if (File.Exists(filePath))
                    {
                        File.SetAttributes(filePath, FileAttributes.Normal);
                    }
                }
            });

            #endregion
            
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