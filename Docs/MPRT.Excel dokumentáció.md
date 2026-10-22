# MPRT.Excel

> **Office-mentes Excel automatizáció UiPath Robotokhoz** — [ClosedXML](https://github.com/ClosedXML/ClosedXML) alapokon

---

## Tartalomjegyzék

- [Célkitűzés](#célkitűzés)
- [Előfeltételek](#előfeltételek)
- [NuGet függőségek](#nuget-függőségek)
- [Architektúra áttekintés](#architektúra-áttekintés)
- [ExcelManager — Használati útmutató](#excelmanager--használati-útmutató)
- [Használati minták](#használati-minták)
  - [Adatolvasás](#adatolvasás)
  - [Adatírás](#adatírás)
  - [Lapkezelés](#lapkezelés)
  - [Formázás](#formázás)
  - [Kombinált / összetett műveletek](#kombinált--összetett-műveletek)
- [Activity referencia](#activity-referencia)
  - [Olvasás (Read)](#olvasás-read)
  - [Írás (Write)](#írás-write)
  - [Lapkezelés (Sheet Management)](#lapkezelés-sheet-management)
  - [Formázás (Format)](#formázás-format)
- [Tervezési elvek](#tervezési-elvek)

---

## Célkitűzés

Az **MPRT.Excel** egy UiPath egyedi activity-könyvtár `.xlsx` fájlok olvasásához és írásához, **Microsoft Office telepítése nélkül** a robot gépen. A nyílt forráskódú ClosedXML könyvtárat tiszta, rétegezett API mögé csomagolja, amelyet a UiPath Studio drag-and-drop activity-kként tesz elérhetővé.

Főbb jellemzők:

- Nincs COM interop, nincs Office/Excel függőség
- Működik headless/szerver környezetekben (UiPath Robot, Orchestrator job runner)
- Minden fájlművelet atomi: megnyitás → módosítás → mentés → bezárás egyetlen `using` blokkban
- A képletek mentéskor automatikusan kiértékelődnek (`EvaluateFormulasBeforeSaving = true`)

---

## Előfeltételek

| Követelmény | Részletek |
|---|---|
| .NET futtatókörnyezet | .NET Framework 4.6.1+ vagy .NET 6+ |
| UiPath Studio | 2022.10+ (activity pack célplatform: Windows / Cross-platform) |
| Microsoft Office | **Nem szükséges** |

---

## NuGet függőségek

| Csomag | Verzió | Szerepe |
|---|---|---|
| `ClosedXML` | ≥ 0.102.x | Fő Excel motor |
| `UiPath.System.Activities` | ≥ 22.x | `CodeActivity`, `InArgument<T>` stb. |
| `DocumentFormat.OpenXml` | ≥ 2.19 | Open XML formátumtámogatás (ClosedXML függőség) |

---

## Architektúra áttekintés

```
MPRT.Excel (logikai réteg)
│
├── Excel.[MetódusNév].cs          — statikus partial osztály; fájlonként egy metódus
├── ExcelManager.cs                — IDisposable burkoló; megnyitás/mentés/zárás életciklus
├── WorksheetExtensions.[Név].cs   — fluent bővítő metódusok IXLWorksheet-hez
└── WorkbookExtensions.[Név].cs    — fluent bővítő metódusok IXLWorkbook-hoz

MPRT.Excel.Activities (UiPath burkolók)
│
├── Read/          — ReadRange, ReadCell, ReadFormula, FindCellValue,
│                    GetSheetNames, SheetExists, GetLastRow, GetLastColumn,
│                    GetColumnLetter
├── Write/         — WriteCell, WriteDataTable, AppendDataTable,
│                    DeleteSheet, RenameSheet, MoveRow, MoveColumn
├── Sheet/         — CopySheet, CopySheetToWorkbook, MoveSheetToFirst,
│                    MoveSheetToLast, SetSheetPosition
└── Format/        — FreezePanes, SortRange, AutoFitColumns,
                     HideRow, UnhideRow, HideColumn, UnhideColumn
```

**Felelősségi körök szétválasztása — szigorú szabály:**

```
Activity  →  Excel.[Metódus] meghívása (statikus, UiPath típusok nélkül)
                ↓
            ExcelManager.Worksheet(név)
                ↓
            WorksheetExtensions.[Metódus](...)  ← minden valódi logika itt van
```

---

## ExcelManager — Használati útmutató

Az `ExcelManager` a könyvtár gerince. Implementálja az `IDisposable` interfészt: a konstruktor megnyitja (vagy létrehozza) a munkafüzetet, a `Dispose()` pedig képletértékeléssel elmenti, majd bezárja.

**Mindig `using` blokkban kell használni.**

### Alap minta

```csharp
using (var excel = new ExcelManager(@"C:\Riportok\kimenet.xlsx"))
{
    var lap = excel.Worksheet("Lap1");   // ha nem létezik, automatikusan létrehozza
    lap.Cell("A1").Value = "Hello World";
}
// A fájl itt mentésre és bezárásra kerül automatikusan
```

### Fluent láncolás

```csharp
using (var excel = new ExcelManager(@"C:\Adatok\riport.xlsx"))
{
    excel.Worksheet("Összefoglaló")
         .WriteDataTable(adatTabla, "A1", includeHeaders: true);

    excel.Worksheet("Összefoglaló")
         .FreezePanes(rowsToFreeze: 1, columnsToFreeze: 0);

    excel.Worksheet("Összefoglaló")
         .AutoFitColumns();
}
```

### Kereszt-munkafüzet másolás (speciális minta)

Ha két különböző fájl között kell lapot másolni, **tilos** két `ExcelManager` példányt egymásba ágyazni. Ehelyett közvetlen `XLWorkbook` + `try/finally` minta használandó — lásd `Excel.CopySheetToWorkbook.cs` mint kanonikus referencia.

```csharp
// A CopySheetToWorkbook belső mintája — csak tájékoztatásul
var forrasWb = new XLWorkbook(forrasFajlUt);
var celWb    = new XLWorkbook(celFajlUt);   // vagy new XLWorkbook() ha új fájl
try
{
    forrasWb.Worksheet(forrasLapNev).CopyTo(celWb, celLapNev);
    celWb.SaveAs(celFajlUt);
}
finally
{
    forrasWb.Dispose();
    celWb.Dispose();
}
```

### Konstruktor validáció

Az `ExcelManager` típusos kivételeket dob, mielőtt a fájlrendszerhez nyúlna:

| Feltétel | Kivétel |
|---|---|
| `filePath` null vagy üres | `ArgumentNullException` |
| A szülőmappa nem létezik | `DirectoryNotFoundException` |
| Az útvonal meglévő mappára mutat | `ArgumentException` |

---

## Használati minták

Az alábbi példák a `MPRT.Excel` core API közvetlen hívásával mutatják be a leggyakoribb RPA-forgatókönyveket. UiPath Studio-ban ezek az activity-k vizuálisan drag-and-drop módon is elvégezhetők — a `Invoke Code` vagy `Execute Macro` activity-ben a core API közvetlenül is hívható.

---

### Adatolvasás

#### Teljes munkalap beolvasása DataTable-be

```csharp
// Has Header = true → az első sor lesz az oszlopnév
// Keep Formatting = false → nyers, típusos értékek (double, DateTime stb.)
DataTable dt = Excel.ReadRange(
    filePath:      @"C:\Adatok\forgalom.xlsx",
    sheetName:     "2024-Q4",
    range:         null,          // teljes használt terület
    hasHeader:     true,
    keepFormatting: false
);
```

#### Meghatározott tartomány olvasása, fejléc nélkül

```csharp
DataTable dt = Excel.ReadRange(
    filePath:  @"C:\Riport\napi.xlsx",
    sheetName: "Raw",
    range:     "B3:F150",
    hasHeader: false
);
```

#### Egyetlen cella értékének kiolvasása

```csharp
// Nyers érték (double, DateTime, bool stb.)
object ertek = Excel.ReadCell(@"C:\Config\beallitasok.xlsx", "Config", "B2");

// Formázott string (pl. "2 024 500 Ft" a számformátum szerint)
object formatalt = Excel.ReadCell(@"C:\Config\beallitasok.xlsx", "Config", "B2",
                                  keepFormatting: true);
```

#### Cella keresése érték alapján

```csharp
// Visszaadja az első egyező cella címét, pl. "C14"
// Ha nem található, üres stringet ad vissza
string cim = Excel.FindCellValue(@"C:\Lista\ugyfelek.xlsx", "Ügyfelek", "Kovács Béla");

if (!string.IsNullOrEmpty(cim))
{
    // A sor sorszámát a cím alapján nyerjük ki
    int sor = int.Parse(System.Text.RegularExpressions.Regex.Match(cim, @"\d+").Value);
}
```

#### Képlet kiolvasása (nem az eredménye, hanem maga a képlet)

```csharp
string keplet = Excel.ReadFormula(@"C:\Riport\osszesito.xlsx", "Összefoglaló", "D10");
// Eredmény pl.: "=SZUM(D2:D9)" vagy "" ha a cellában nincs képlet
```

#### Munkalapok listázása + létezés ellenőrzése

```csharp
// Az összes lap neve a munkafüzet sorrendjében
List<string> lapok = Excel.GetSheetNames(@"C:\Adatok\munkafuzet.xlsx");
// → ["Január", "Február", "Március", "Összefoglaló"]

// Feltételes feldolgozás: csak akkor olvasunk, ha a lap létezik
bool letezik = Excel.SheetExists(@"C:\Adatok\munkafuzet.xlsx", "Összefoglaló");
if (letezik)
{
    var osszefoglalo = Excel.ReadRange(@"C:\Adatok\munkafuzet.xlsx", "Összefoglaló");
}
```

#### Dinamikus tartomány meghatározása (utolsó sor / oszlop)

```csharp
// Az "A" oszlop utolsó adatot tartalmazó sora
int utolsoSor = Excel.GetLastRow(@"C:\Adatok\log.xlsx", "Napló", "A");
// → pl. 1523

// Az 1. sor utolsó adatot tartalmazó oszlopa (fejléc)
string utolsoOszlop = Excel.GetLastColumn(@"C:\Adatok\log.xlsx", "Napló", rowNumber: 1);
// → pl. "K"

// Dinamikus tartomány összerakása
string tartomany = $"A1:{utolsoOszlop}{utolsoSor}";  // "A1:K1523"
DataTable dt = Excel.ReadRange(@"C:\Adatok\log.xlsx", "Napló", tartomany);
```

#### Oszlop keresése fejléc neve alapján

```csharp
// Megadja a "Nettó összeg" fejlécű oszlop betűjelét (pl. "F")
string oszlopBetu = Excel.GetColumnLetter(
    filePath:   @"C:\Szamla\szamlak.xlsx",
    sheetName:  "Adatok",
    columnName: "Nettó összeg",
    headerRow:  1
);
string tartomany = $"{oszlopBetu}2:{oszlopBetu}{utolsoSor}";
```

---

### Adatírás

#### Egyetlen cella írása (különböző típusok)

```csharp
string fajl    = @"C:\Riport\kimenet.xlsx";
string lap     = "Összefoglaló";

Excel.WriteCell(fajl, lap, "A1", "Feldolgozás dátuma:");
Excel.WriteCell(fajl, lap, "B1", DateTime.Today);        // DateTime
Excel.WriteCell(fajl, lap, "A2", "Rekordszám:");
Excel.WriteCell(fajl, lap, "B2", 1523);                  // int
Excel.WriteCell(fajl, lap, "A3", "Sikeres:");
Excel.WriteCell(fajl, lap, "B3", true);                  // bool
Excel.WriteCell(fajl, lap, "B4", "=SZUM(B5:B100)");     // képlet stringként
```

> **Megjegyzés:** Ha több cellát kell egyszerre írni, érdemes egyetlen `ExcelManager` blokkban elvégezni, hogy a fájl csak egyszer nyíljon meg és egyszer mentődjön.

```csharp
using (var excel = new ExcelManager(@"C:\Riport\kimenet.xlsx"))
{
    var lap = excel.Worksheet("Összefoglaló");
    lap.Cell("A1").Value = "Feldolgozás dátuma:";
    lap.Cell("B1").Value = DateTime.Today;
    lap.Cell("A2").Value = "Rekordszám:";
    lap.Cell("B2").Value = 1523;
}
```

#### DataTable teljes kiírása fejléccel

```csharp
// A DataTable oszlopnevei lesznek az Excel fejlécek
Excel.WriteDataTable(
    filePath:       @"C:\Export\eredmeny.xlsx",
    sheetName:      "Adatok",
    dataTable:      feldolgozottAdatok,
    startCell:      "A1",
    includeHeaders: true
);
```

#### DataTable hozzáfűzése meglévő adatok után

```csharp
// Nem kell tudni, hány sor van már a lapon — automatikusan az utolsó sor alá ír
// A fejlécet NEM írja ki újra
Excel.AppendDataTable(
    filePath:  @"C:\Log\feldolgozas_log.xlsx",
    sheetName: "Log",
    dataTable: ujSorok
);
```

#### Munkalap törlése és átnevezése

```csharp
// Ideiglenes "Munka" lap törlése az elkészült riportból
Excel.DeleteSheet(@"C:\Riport\vegso.xlsx", "Munka");

// Dinamikusan generált névből végleges névre
Excel.RenameSheet(@"C:\Riport\vegso.xlsx", "Sheet1", $"Riport_{DateTime.Today:yyyy-MM}");
```

#### Sor és oszlop mozgatása

```csharp
// A 3. sort a 10. pozícióba helyezi (pl. prioritásos sor kiemelése)
Excel.MoveRow(@"C:\Feladatok\lista.xlsx", "Feladatok", sourceRowNumber: 3, destinationRowNumber: 10);

// A 4. oszlopot (pl. "D") a 2. pozícióba helyezi
Excel.MoveColumn(@"C:\Adatok\tabla.xlsx", "Adatok", sourceColumnNumber: 4, destinationColumnNumber: 2);
```

---

### Lapkezelés

#### Munkalap másolása ugyanazon a fájlon belül (mentési sablon)

```csharp
// A "Sablon" lapot lemásolja "2024-December" névvel
Excel.CopySheet(
    filePath:        @"C:\Havi\forgalom.xlsx",
    sourceSheetName: "Sablon",
    newSheetName:    $"{DateTime.Today:yyyy-MMMM}"
);
```

#### Munkalap másolása másik fájlba

```csharp
// A "Összefoglaló" lap átkerül a célfájlba; ha az nem létezik, létrejön
Excel.CopySheetToWorkbook(
    sourceFilePath:      @"C:\Forras\forrás_adat.xlsx",
    sourceSheetName:     "Összefoglaló",
    destinationFilePath: @"C:\Cel\konszolidalt.xlsx",
    destinationSheetName: "Q4-Összefoglaló"
);
```

#### Lapok rendezése pozíció szerint

```csharp
string fajl = @"C:\Riport\havi_riport.xlsx";

// Az "Összefoglaló" lap kerüljön előre
Excel.MoveSheetToFirst(fajl, "Összefoglaló");

// A "Napló" lap kerüljön hátulra
Excel.MoveSheetToLast(fajl, "Napló");

// A "Január" lap kerüljön a 2. pozícióba
Excel.SetSheetPosition(fajl, "Január", position: 2);
```

#### Lapok feldolgozása iterálva

```csharp
// Minden lapot feldolgoz, kivéve a "Sablon" nevűt
List<string> lapok = Excel.GetSheetNames(@"C:\Adatok\munkafuzet.xlsx");

foreach (string lapNev in lapok.Where(n => n != "Sablon"))
{
    DataTable dt = Excel.ReadRange(@"C:\Adatok\munkafuzet.xlsx", lapNev);
    // ... feldolgozás
}
```

---

### Formázás

#### Fejlécsor rögzítése (Freeze Panes)

```csharp
// Csak az első sor rögzítése (leggyakoribb eset)
Excel.FreezePanes(@"C:\Riport\nagy_tabla.xlsx", "Adatok", rowsToFreeze: 1, columnsToFreeze: 0);

// Első sor ÉS első két oszlop rögzítése (pl. azonosítók láthatóak maradnak görgetéskor)
Excel.FreezePanes(@"C:\Riport\nagy_tabla.xlsx", "Adatok", rowsToFreeze: 1, columnsToFreeze: 2);
```

#### Tartomány rendezése

```csharp
string fajl = @"C:\Adatok\ugyfelek.xlsx";
string lap  = "Ügyfelek";
string tar  = "A1:E500";

// Rendezés oszlopszám szerint (2. oszlop = B → Vezetéknév), csökkenő
Excel.SortRange(fajl, lap, tar, columnNumber: 2, XLSortOrder.Descending, hasHeader: true);

// Rendezés oszlopbetű szerint
Excel.SortRange(fajl, lap, tar, columnLetter: "C", XLSortOrder.Ascending, hasHeader: true);

// Rendezés fejléc neve szerint (rugalmas, ha az oszlopok sorrendje változhat)
Excel.SortRangeByName(fajl, lap, tar, columnName: "Utolsó vásárlás", XLSortOrder.Descending);
```

#### Oszlopszélesség igazítása

```csharp
// Az összes oszlop igazítása a tartalomhoz
Excel.AutoFitColumns(@"C:\Export\kimenet.xlsx", "Adatok");

// Csak az 1., 3. és 5. oszlop igazítása
Excel.AutoFitColumns(@"C:\Export\kimenet.xlsx", "Adatok", 1, 3, 5);
```

#### Sorok és oszlopok elrejtése / megjelenítése

```csharp
string fajl = @"C:\Riport\belso_riport.xlsx";
string lap  = "Részletek";

// Belső kalkulációs sorok elrejtése (pl. 2–5. sor)
for (int i = 2; i <= 5; i++)
    Excel.HideRow(fajl, lap, i);

// Segédoszlopok elrejtése (pl. 6. és 7. oszlop = F, G)
Excel.HideColumn(fajl, lap, 6);
Excel.HideColumn(fajl, lap, 7);

// Audit során az összes visszaállítható
Excel.UnhideRow(fajl, lap, 2);
Excel.UnhideColumn(fajl, lap, 6);
```

---

### Kombinált / összetett műveletek

#### Napi riport generálása (teljes folyamat)

```csharp
string sablon   = @"C:\Sablonok\napi_riport_sablon.xlsx";
string kimenet  = $@"C:\Riportok\{DateTime.Today:yyyy-MM-dd}_riport.xlsx";

// 1. Sablon másolása → ne a sablont módosítsuk
Excel.CopySheetToWorkbook(sablon, "Adatok", kimenet, "Adatok");
Excel.CopySheetToWorkbook(sablon, "Összefoglaló", kimenet, "Összefoglaló");

// 2. Adatok kiírása
Excel.WriteDataTable(kimenet, "Adatok", feldolgozottDt, "A1", includeHeaders: true);

// 3. Fejléc metaadatok
Excel.WriteCell(kimenet, "Összefoglaló", "B1", DateTime.Today);
Excel.WriteCell(kimenet, "Összefoglaló", "B2", feldolgozottDt.Rows.Count);

// 4. Formázás egyetlen using blokkban (egyetlen fájlmegnyitás)
using (var excel = new ExcelManager(kimenet))
{
    var adatLap = excel.Worksheet("Adatok");
    adatLap.FreezePanes(1, 0);
    adatLap.AutoFitColumns();

    var ossz = excel.Worksheet("Összefoglaló");
    ossz.FreezePanes(1, 0);
}

// 5. Rendezés (Számla összeg csökkenő)
int utolsoSor = Excel.GetLastRow(kimenet, "Adatok", "A");
Excel.SortRangeByName(kimenet, "Adatok", $"A1:Z{utolsoSor}", "Számla összeg", XLSortOrder.Descending);

// 6. Összefoglaló lap előre
Excel.MoveSheetToFirst(kimenet, "Összefoglaló");
```

#### Log fájl kezelése (folyamatos hozzáfűzés)

```csharp
string logFajl = @"C:\Log\robot_log.xlsx";
string logLap  = "Log";

// Ha a fájl még nem létezik, az ExcelManager automatikusan létrehozza
// Ha létezik, az AppendDataTable az utolsó sor alá ír
var ujBejegyzesek = new DataTable();
ujBejegyzesek.Columns.Add("Időbélyeg", typeof(DateTime));
ujBejegyzesek.Columns.Add("Tranzakció ID", typeof(string));
ujBejegyzesek.Columns.Add("Státusz", typeof(string));
ujBejegyzesek.Rows.Add(DateTime.Now, "TXN-00123", "Sikeres");

// Első futáskor fejlécet is ír (WriteDataTable), utána csak hozzáfűz (AppendDataTable)
if (!System.IO.File.Exists(logFajl) || Excel.GetLastRow(logFajl, logLap, "A") == 0)
    Excel.WriteDataTable(logFajl, logLap, ujBejegyzesek, "A1", includeHeaders: true);
else
    Excel.AppendDataTable(logFajl, logLap, ujBejegyzesek);
```

#### Több forrás konszolidálása egyetlen fájlba

```csharp
string[] forrasFajlok = Directory.GetFiles(@"C:\Havi\", "*.xlsx");
string celFajl        = @"C:\Konszolidalt\osszes_adat.xlsx";

foreach (string forras in forrasFajlok)
{
    string lapNev = Path.GetFileNameWithoutExtension(forras);  // pl. "2024-Január"

    // Minden forrás fájlból átmásoljuk az "Adatok" lapot saját névvel
    Excel.CopySheetToWorkbook(forras, "Adatok", celFajl, lapNev);
}

// Az "Összefoglaló" lap előre, a többi ábécé sorrendbe
Excel.MoveSheetToFirst(celFajl, "Összefoglaló");
```

#### Sablon-alapú cellafrissítés (érték keresés + írás)

```csharp
string fajl = @"C:\Sablon\szerzodes_sablon.xlsx";
string lap  = "Adatok";

// Megkeressük a "[[ÜGYFÉLNÉV]]" placeholder celláját, majd felülírjuk
string cim = Excel.FindCellValue(fajl, lap, "[[ÜGYFÉLNÉV]]");
if (!string.IsNullOrEmpty(cim))
    Excel.WriteCell(fajl, lap, cim, "Kovács Bt.");

cim = Excel.FindCellValue(fajl, lap, "[[DÁTUM]]");
if (!string.IsNullOrEmpty(cim))
    Excel.WriteCell(fajl, lap, cim, DateTime.Today);

cim = Excel.FindCellValue(fajl, lap, "[[ÖSSZEG]]");
if (!string.IsNullOrEmpty(cim))
    Excel.WriteCell(fajl, lap, cim, 1_250_000.0);
```

---

## Activity referencia

> **Minden activity-n közös bemeneti tulajdonságok** (az egyedi táblázatokban nem ismétlődnek):
> - **1. File Path** `[Kötelező]` — az `.xlsx` fájl teljes elérési útja
> - **2. Sheet Name** `[Kötelező]` — munkalap neve (ha nem létezik, automatikusan létrejön)

---

### Olvasás (Read)

| Megjelenített név | Bemeneti tulajdonságok | Kimeneti tulajdonságok | Megjegyzés |
|---|---|---|---|
| **Read Range** | 3. Range *(opcionális)*, Has Header *(bool, alap: true)*, Keep Formatting *(bool, alap: false)* | Data Table `DataTable` | Üres Range → teljes használt területet olvassa. Egyetlen cellacím → szomszédos adatrégiót bővíti. |
| **Read Cell** | 3. Cell Address, Keep Formatting *(bool, alap: false)* | Value `object` | Nyers típusos értéket vagy formázott stringet ad vissza. |
| **Read Formula** | 3. Cell Address | Formula `string` | A cellában lévő képletet adja vissza stringként (pl. `"=SZUM(A1:A10)"`). Üres string, ha nincs képlet. |
| **Find Cell Value** | 3. Value to Find `object` | Cell Address `string` | Az első egyező cella címét adja vissza (pl. `"B7"`). Nem találás esetén üres string. |
| **Get Sheet Names** | *(csak File Path)* | Sheet Names `List<string>` | Az összes munkalap nevét visszaadja a munkafüzet-sorrendben. |
| **Sheet Exists** | 3. Sheet Name | Exists `bool` | `true`, ha a megadott nevű munkalap létezik. |
| **Get Last Row In Column** | 3. Column Letter *(pl. "A")* | Last Row Number `int` | Az adott oszlop utolsó nem üres cellájának sorszámát adja vissza. |
| **Get Last Column Letter In Row** | 3. Row Number `int` | Column Letter `string` | Az adott sor utolsó nem üres cellájának oszlopbetűjelét adja vissza (pl. `"D"`). Üres sor esetén üres string. |
| **Get Letter Of Specific Column** | 3. Column Name, 4. Header Row *(int, alap: 1)* | Column Letter `string` | Fejléc szövege alapján keresi meg az oszlopot, és visszaadja a betűjelét (pl. `"AB"`). |

---

### Írás (Write)

| Megjelenített név | Bemeneti tulajdonságok | Kimeneti tulajdonságok | Megjegyzés |
|---|---|---|---|
| **Write Cell** | 3. Cell Address, 4. Value `object` | — | Bármilyen szerializálható értéket fogad (string, int, double, DateTime, bool). |
| **Write DataTable** | 3. Data Table `DataTable`, 4. Start Cell *(alap: "A1")*, 5. Include Headers *(bool, alap: true)* | — | A kezdőcellától felülírja az adatokat. `InsertTable` / `InsertData` használatával tömeges teljesítmény. |
| **Append DataTable** | 3. Data Table `DataTable` | — | Megkeresi az utolsó használt sort, és alá írja az adatokat. A fejlécsor soha nem duplikálódik. |
| **Delete Sheet** | *(File Path + Sheet Name)* | — | Véglegesen törli a munkalapot. A törlés után az objektum érvénytelen lesz. |
| **Rename Sheet** | 3. New Sheet Name | — | A munkalapot helyben átnevezi. |
| **Move Row** | 3. Source Row Number `int`, 4. Destination Row Number `int` | — | Soron belüli mozgatás ugyanazon a munkalapon. |
| **Move Column** | 3. Source Column Number `int`, 4. Destination Column Number `int` | — | Oszlopon belüli mozgatás ugyanazon a munkalapon. |

---

### Lapkezelés (Sheet Management)

| Megjelenített név | Bemeneti tulajdonságok | Kimeneti tulajdonságok | Megjegyzés |
|---|---|---|---|
| **Copy Sheet** | 3. New Sheet Name *(opcionális)* | — | Munkalapot másol **ugyanazon** a fájlon belül. |
| **Copy Sheet To Workbook** | 1. Source File Path `[Kötelező]`, 2. Source Sheet Name `[Kötelező]`, 3. Destination File Path, 4. Destination Sheet Name | — | Fájlok közötti másolás. Ha a célfájl nem létezik, automatikusan létrejön. Speciális duális-munkafüzet mintát használ (nem `ExcelManager`). |
| **Move Sheet To First** | *(File Path + Sheet Name)* | — | A lapot az 1-es pozícióba helyezi. |
| **Move Sheet To Last** | *(File Path + Sheet Name)* | — | A lapot az összes lap számával egyező utolsó pozícióba helyezi. |
| **Set Sheet Position** | 3. Position `int` *(1-alapú)* | — | Explicit 1-alapú pozícióindex beállítása. |

---

### Formázás (Format)

| Megjelenített név | Bemeneti tulajdonságok | Kimeneti tulajdonságok | Megjegyzés |
|---|---|---|---|
| **Freeze Panes** | 3. Rows to Freeze `int` *(alap: 1)*, 4. Columns to Freeze `int` *(alap: 0)* | — | Mindkét érték lehet 0 (nincs rögzítés abban a tengelyben). |
| **Sort Range** | 3. Range *(pl. "A1:D100")*, 4a. Sort By Column Number `int?`, 4b. Sort By Column Letter `string`, 4c. Sort By Column Name `string`, 5. Has Header *(bool, alap: true)*, 6. Sort Order *(Ascending / Descending)* | — | A 4a / 4b / 4c közül pontosan **egyet** kell megadni. A Column Name változat a fejlécsorban keres névazonosság alapján. |
| **AutoFit Columns** | 3. Columns `int[]` *(opcionális)* | — | Konkrét oszlopszámok megadásával csak azok igazítódnak; üresen hagyva az összes használt oszlop igazítódik. |
| **Hide Row** | 3. Row Number `int` | — | Elrejti a sort. A sorszámnak ≥ 1-nek kell lennie. |
| **Unhide Row** | 3. Row Number `int` | — | Egy rejtett sort láthatóvá tesz. |
| **Hide Column** | 3. Column Number `int` | — | Elrejti az oszlopot. Az oszlopszámnak ≥ 1-nek kell lennie. |
| **Unhide Column** | 3. Column Number `int` | — | Egy rejtett oszlopot láthatóvá tesz. |

---

## Tervezési elvek

### Réteg-felelősségek

| Réteg | Megengedett | Tilos |
|---|---|---|
| `Activity` osztály | `InArgument.Get()`, `OutArgument.Set()`, `Execute(context)` | Üzleti logika, ClosedXML típusok, ciklusok |
| `Excel.[Metódus]` | `using (var m = new ExcelManager(...))` + delegálás az extension-re | `CodeActivityContext`, `InArgument` |
| `WorksheetExtensions` | Tiszta ClosedXML műveletek, LINQ egysoros kifejezések | Fájl I/O, UiPath típusok |

### Mentési viselkedés

Minden `ExcelManager.Dispose()` hívás az alábbi beállítással ment:

```csharp
_workbook.SaveAs(_filePath, new SaveOptions { EvaluateFormulasBeforeSaving = true });
```

Ez azt jelenti, hogy minden mentéskor az összes képlet újraszámolódik — a fájl mindig aktuális, kiszámított értékeket tartalmaz, amikor Excelben megnyitják.

### Bemeneti validáció

Minden bemenet a lehető legkorábban validálódik, típusos kivételekkel és leíró magyar hibaüzenetekkel. A validáció az `ExcelManager` konstruktorában (fájlútvonal) és az egyes extension metódusokban (null munkalap, érvénytelen sor-/oszlopszám) történik.

### Partial class / egy-fájl-egy-metódus konvenció

```
MPRT.Excel/
  Excel.ReadRange.cs
  Excel.WriteCell.cs
  Excel.FreezePanes.cs
  ...
  WorksheetExtension.ReadRange.cs
  WorksheetExtension.WriteDataTable.cs
  ...
```

Minden fájl pontosan egy metódust tartalmaz. A fájlnév megegyezik a metódus nevével. Ez minimalizálja a diff-eket és triviálissá teszi a navigációt.

---

*MPRT.Excel — ClosedXML alapokon · Office nem szükséges*
