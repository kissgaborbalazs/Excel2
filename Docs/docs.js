const DOCS_DATA = {
  "_info": "Dokumentációs adatfájl. Új activity hozzáadásához csak az activityGroups tömböt kell bővíteni — kódolás nem szükséges. Másik library dokumentálásához cseréld ki az egész fájlt.",
  "_schema": {
    "meta": "Library fejléc adatok",
    "highlights": "Főbb jellemzők kártyák (icon, title, description)",
    "prerequisites": "Előfeltételek táblázat (requirement, details)",
    "nuget": "NuGet függőségek (package, version, role)",
    "architecture": "Architektúra diagram és rétegek",
    "excelManager": "Core osztály leírás, hibák, kód példák",
    "commonInputs": "Minden activity-n közös bemenetek",
    "activityGroups": "Activity csoportok. Új group: id, title, titleEn, activities tömb. Új activity: displayName, inputs[], outputs[], notes",
    "designPrinciples": "Tervezési elvek (title, content, code?)"
  },
  "meta": {
    "name": "MPRT.Excel",
    "tagline": "Office-mentes Excel automatizáció UiPath Robotokhoz",
    "engine": "ClosedXML",
    "engineUrl": "https://github.com/ClosedXML/ClosedXML",
    "version": "1.0.0",
    "description": "Egyedi UiPath activity-könyvtár .xlsx fájlok olvasásához és írásához, Microsoft Office telepítése nélkül."
  },
  "highlights": [
    {
      "icon": "🚫",
      "title": "Office-mentes",
      "description": "Nincs COM interop, nincs Excel/Office függőség. Működik headless szerveren, UiPath Robot gépen telepítés nélkül."
    },
    {
      "icon": "⚡",
      "title": "Atomi fájlkezelés",
      "description": "Megnyitás → módosítás → mentés → bezárás egyetlen using blokkban. Soha nem marad nyitva fájl."
    },
    {
      "icon": "🔗",
      "title": "Fluent API",
      "description": "Láncolható extension metódusok IXLWorksheet-en és IXLWorkbook-on. LINQ-alapú, tömör egysoros kifejezések."
    },
    {
      "icon": "🧮",
      "title": "Képletkiértékelés",
      "description": "Dispose() mindig EvaluateFormulasBeforeSaving = true beállítással ment — a fájl mindig aktuális értékeket tartalmaz."
    }
  ],
  "prerequisites": [
    {
      "requirement": ".NET futtatókörnyezet",
      "details": ".NET Framework 4.6.1+ vagy .NET 6+"
    },
    {
      "requirement": "UiPath Studio",
      "details": "2022.10+ (activity pack célplatform: Windows / Cross-platform)"
    },
    {
      "requirement": "Microsoft Office",
      "details": "Nem szükséges ✓"
    }
  ],
  "nuget": [
    {
      "package": "ClosedXML",
      "version": "≥ 0.102.x",
      "role": "Fő Excel motor"
    },
    {
      "package": "UiPath.System.Activities",
      "version": "≥ 22.x",
      "role": "CodeActivity, InArgument<T> stb."
    },
    {
      "package": "DocumentFormat.OpenXml",
      "version": "≥ 2.19",
      "role": "Open XML formátumtámogatás (ClosedXML függőség)"
    }
  ],
  "architecture": {
    "diagram": [
      "MPRT.Excel  (logikai réteg)",
      "│",
      "├── Excel.[MetódusNév].cs          — statikus partial osztály; fájlonként egy metódus",
      "├── ExcelManager.cs                — IDisposable burkoló; megnyitás/mentés/zárás életciklus",
      "├── WorksheetExtensions.[Név].cs   — fluent bővítő metódusok IXLWorksheet-hez",
      "└── WorkbookExtensions.[Név].cs    — fluent bővítő metódusok IXLWorkbook-hoz",
      "",
      "MPRT.Excel.Activities  (UiPath burkolók)",
      "│",
      "├── Read/    — ReadRange, ReadCell, ReadFormula, FindCellValue,",
      "│              GetSheetNames, SheetExists, GetLastRow, GetLastColumn, GetColumnLetter",
      "├── Write/   — WriteCell, WriteDataTable, AppendDataTable,",
      "│              DeleteSheet, RenameSheet, MoveRow, MoveColumn",
      "├── Sheet/   — CopySheet, CopySheetToWorkbook, MoveSheetToFirst,",
      "│              MoveSheetToLast, SetSheetPosition",
      "└── Format/  — FreezePanes, SortRange, AutoFitColumns,",
      "               HideRow, UnhideRow, HideColumn, UnhideColumn"
    ],
    "flow": [
      {
        "step": "Activity",
        "desc": "InArgument.Get() + delegálás",
        "arrow": "→"
      },
      {
        "step": "Excel.[Metódus]",
        "desc": "using ExcelManager + delegálás",
        "arrow": "→"
      },
      {
        "step": "WorksheetExtensions",
        "desc": "Tiszta ClosedXML logika",
        "arrow": ""
      }
    ],
    "layers": [
      {
        "layer": "Activity osztály",
        "allowed": "InArgument.Get(), OutArgument.Set(), Execute(context)",
        "forbidden": "Üzleti logika, ClosedXML típusok, ciklusok"
      },
      {
        "layer": "Excel.[Metódus]",
        "allowed": "using (var m = new ExcelManager(...)) + delegálás az extension-re",
        "forbidden": "CodeActivityContext, InArgument"
      },
      {
        "layer": "WorksheetExtensions",
        "allowed": "Tiszta ClosedXML műveletek, LINQ egysoros kifejezések",
        "forbidden": "Fájl I/O, UiPath típusok"
      }
    ]
  },
  "excelManager": {
    "description": "Az ExcelManager a könyvtár gerince. Implementálja az IDisposable interfészt: a konstruktor megnyitja (vagy létrehozza) a munkafüzetet, a Dispose() pedig képletértékeléssel elmenti, majd bezárja. Mindig using blokkban kell használni.",
    "constructorErrors": [
      {
        "condition": "filePath null vagy üres",
        "exception": "ArgumentNullException"
      },
      {
        "condition": "A szülőmappa nem létezik",
        "exception": "DirectoryNotFoundException"
      },
      {
        "condition": "Az útvonal meglévő mappára mutat",
        "exception": "ArgumentException"
      }
    ],
    "codeExamples": [
      {
        "title": "Alap minta",
        "lang": "csharp",
        "code": [
          "using (var excel = new ExcelManager(@\"C:\\Riportok\\kimenet.xlsx\"))",
          "{",
          "    var lap = excel.Worksheet(\"Lap1\");   // ha nem létezik, automatikusan létrehozza",
          "    lap.Cell(\"A1\").Value = \"Hello World\";",
          "}",
          "// A fájl itt mentésre és bezárásra kerül automatikusan"
        ]
      },
      {
        "title": "Fluent láncolás — több művelet egyetlen megnyitással",
        "lang": "csharp",
        "code": [
          "using (var excel = new ExcelManager(@\"C:\\Adatok\\riport.xlsx\"))",
          "{",
          "    excel.Worksheet(\"Összefoglaló\")",
          "         .WriteDataTable(adatTabla, \"A1\", includeHeaders: true);",
          "",
          "    excel.Worksheet(\"Összefoglaló\")",
          "         .FreezePanes(rowsToFreeze: 1, columnsToFreeze: 0);",
          "",
          "    excel.Worksheet(\"Összefoglaló\")",
          "         .AutoFitColumns();",
          "}"
        ]
      },
      {
        "title": "Kereszt-munkafüzet másolás (speciális minta)",
        "lang": "csharp",
        "code": [
          "// TILOS két ExcelManager-t egymásba ágyazni!",
          "// Helyette: közvetlen XLWorkbook + try/finally",
          "var forrasWb = new XLWorkbook(forrasFajlUt);",
          "var celWb    = new XLWorkbook(celFajlUt);",
          "try",
          "{",
          "    forrasWb.Worksheet(forrasLapNev).CopyTo(celWb, celLapNev);",
          "    celWb.SaveAs(celFajlUt);",
          "}",
          "finally",
          "{",
          "    forrasWb.Dispose();",
          "    celWb.Dispose();",
          "}"
        ]
      }
    ]
  },
  "commonInputs": [
    {
      "name": "1. File Path",
      "required": true,
      "type": "string",
      "description": "Az .xlsx fájl teljes elérési útja"
    },
    {
      "name": "2. Sheet Name",
      "required": true,
      "type": "string",
      "description": "Munkalap neve (ha nem létezik, automatikusan létrejön)"
    }
  ],
  "activityGroups": [
    {
      "id": "read",
      "title": "Olvasás",
      "titleEn": "Read",
      "activities": [
        {
          "displayName": "Read Range",
          "inputs": [
            {
              "name": "3. Range",
              "required": false,
              "type": "string",
              "default": null,
              "description": "Pl. \"B3:F150\". Üres → teljes használt terület. Egyetlen cím → szomszédos régió."
            },
            {
              "name": "Has Header",
              "required": false,
              "type": "bool",
              "default": "true",
              "description": "Az első sor oszlopnévként kerül be a DataTable-be."
            },
            {
              "name": "Keep Formatting",
              "required": false,
              "type": "bool",
              "default": "false",
              "description": "true → formázott string; false → nyers típusos érték."
            }
          ],
          "outputs": [
            {
              "name": "Data Table",
              "type": "DataTable"
            }
          ],
          "notes": "Egyetlen cellacím Range esetén a ClosedXML a szomszédos adatrégiót bővíti automatikusan."
        },
        {
          "displayName": "Read Cell",
          "inputs": [
            {
              "name": "3. Cell Address",
              "required": true,
              "type": "string",
              "default": null,
              "description": "Pl. \"B7\""
            },
            {
              "name": "Keep Formatting",
              "required": false,
              "type": "bool",
              "default": "false",
              "description": "true → formázott string (pl. \"2 024 500 Ft\"); false → nyers érték."
            }
          ],
          "outputs": [
            {
              "name": "Value",
              "type": "object"
            }
          ],
          "notes": "Nyers típusos értéket (double, DateTime, bool stb.) vagy formázott stringet ad vissza."
        },
        {
          "displayName": "Read Formula",
          "inputs": [
            {
              "name": "3. Cell Address",
              "required": true,
              "type": "string",
              "default": null,
              "description": "Pl. \"D10\""
            }
          ],
          "outputs": [
            {
              "name": "Formula",
              "type": "string"
            }
          ],
          "notes": "A cellában lévő képletet adja vissza stringként (pl. \"=SZUM(A1:A10)\"). Üres string, ha nincs képlet."
        },
        {
          "displayName": "Find Cell Value",
          "inputs": [
            {
              "name": "3. Value to Find",
              "required": true,
              "type": "object",
              "default": null,
              "description": "Keresett érték (bármilyen típus)"
            }
          ],
          "outputs": [
            {
              "name": "Cell Address",
              "type": "string"
            }
          ],
          "notes": "Az első egyező cella címét adja vissza (pl. \"C14\"). Nem találás esetén üres string."
        },
        {
          "displayName": "Get Sheet Names",
          "inputs": [],
          "outputs": [
            {
              "name": "Sheet Names",
              "type": "List&lt;string&gt;"
            }
          ],
          "notes": "Csak File Path szükséges (Sheet Name elhagyható). Az összes munkalap nevét visszaadja a munkafüzet-sorrendben."
        },
        {
          "displayName": "Sheet Exists",
          "inputs": [
            {
              "name": "3. Sheet Name",
              "required": true,
              "type": "string",
              "default": null,
              "description": "Ellenőrzendő munkalap neve"
            }
          ],
          "outputs": [
            {
              "name": "Exists",
              "type": "bool"
            }
          ],
          "notes": "true, ha a megadott nevű munkalap létezik a munkafüzetben."
        },
        {
          "displayName": "Get Last Row In Column",
          "inputs": [
            {
              "name": "3. Column Letter",
              "required": true,
              "type": "string",
              "default": null,
              "description": "Pl. \"A\""
            }
          ],
          "outputs": [
            {
              "name": "Last Row Number",
              "type": "int"
            }
          ],
          "notes": "Az adott oszlop utolsó nem üres cellájának sorszámát adja vissza. Dinamikus tartomány meghatározáshoz hasznos."
        },
        {
          "displayName": "Get Last Column Letter In Row",
          "inputs": [
            {
              "name": "3. Row Number",
              "required": true,
              "type": "int",
              "default": null,
              "description": "Pl. 1 (fejlécsor)"
            }
          ],
          "outputs": [
            {
              "name": "Column Letter",
              "type": "string"
            }
          ],
          "notes": "Az adott sor utolsó nem üres cellájának oszlopbetűjelét adja vissza (pl. \"K\"). Üres sor esetén üres string."
        },
        {
          "displayName": "Get Letter Of Specific Column",
          "inputs": [
            {
              "name": "3. Column Name",
              "required": true,
              "type": "string",
              "default": null,
              "description": "Fejléc szövege, pl. \"Nettó összeg\""
            },
            {
              "name": "4. Header Row",
              "required": false,
              "type": "int",
              "default": "1",
              "description": "Fejlécsor sorszáma"
            }
          ],
          "outputs": [
            {
              "name": "Column Letter",
              "type": "string"
            }
          ],
          "notes": "Fejléc szövege alapján keresi meg az oszlopot, és visszaadja a betűjelét (pl. \"AB\")."
        }
      ]
    },
    {
      "id": "write",
      "title": "Írás",
      "titleEn": "Write",
      "activities": [
        {
          "displayName": "Write Cell",
          "inputs": [
            {
              "name": "3. Cell Address",
              "required": true,
              "type": "string",
              "default": null,
              "description": "Pl. \"B7\""
            },
            {
              "name": "4. Value",
              "required": true,
              "type": "object",
              "default": null,
              "description": "string, int, double, DateTime, bool, képlet string"
            }
          ],
          "outputs": [],
          "notes": "Bármilyen szerializálható értéket fogad. Képlet stringként adható meg (pl. \"=SZUM(B5:B100)\")."
        },
        {
          "displayName": "Write DataTable",
          "inputs": [
            {
              "name": "3. Data Table",
              "required": true,
              "type": "DataTable",
              "default": null,
              "description": "Forrás DataTable"
            },
            {
              "name": "4. Start Cell",
              "required": false,
              "type": "string",
              "default": "\"A1\"",
              "description": "Bal felső sarokcella"
            },
            {
              "name": "5. Include Headers",
              "required": false,
              "type": "bool",
              "default": "true",
              "description": "Oszlopneveket fejlécként írja ki"
            }
          ],
          "outputs": [],
          "notes": "A kezdőcellától felülírja az adatokat. Tömeges teljesítmény InsertTable/InsertData használatával."
        },
        {
          "displayName": "Append DataTable",
          "inputs": [
            {
              "name": "3. Data Table",
              "required": true,
              "type": "DataTable",
              "default": null,
              "description": "Hozzáfűzendő adatok"
            }
          ],
          "outputs": [],
          "notes": "Megkeresi az utolsó használt sort, és alá írja az adatokat. A fejlécsor soha nem duplikálódik."
        },
        {
          "displayName": "Delete Sheet",
          "inputs": [],
          "outputs": [],
          "notes": "Véglegesen törli a munkalapot. Csak File Path + Sheet Name szükséges."
        },
        {
          "displayName": "Rename Sheet",
          "inputs": [
            {
              "name": "3. New Sheet Name",
              "required": true,
              "type": "string",
              "default": null,
              "description": "Új munkalapnév"
            }
          ],
          "outputs": [],
          "notes": "A munkalapot helyben átnevezi. Az eredeti lapra mutató hivatkozások a munkafüzetben frissülnek."
        },
        {
          "displayName": "Move Row",
          "inputs": [
            {
              "name": "3. Source Row Number",
              "required": true,
              "type": "int",
              "default": null,
              "description": "Forrás sor sorszáma (≥ 1)"
            },
            {
              "name": "4. Destination Row Number",
              "required": true,
              "type": "int",
              "default": null,
              "description": "Cél sor sorszáma (≥ 1)"
            }
          ],
          "outputs": [],
          "notes": "Sor mozgatása ugyanazon a munkalapon belül."
        },
        {
          "displayName": "Move Column",
          "inputs": [
            {
              "name": "3. Source Column Number",
              "required": true,
              "type": "int",
              "default": null,
              "description": "Forrás oszlopszám (≥ 1)"
            },
            {
              "name": "4. Destination Column Number",
              "required": true,
              "type": "int",
              "default": null,
              "description": "Cél oszlopszám (≥ 1)"
            }
          ],
          "outputs": [],
          "notes": "Oszlop mozgatása ugyanazon a munkalapon belül."
        }
      ]
    },
    {
      "id": "sheet",
      "title": "Lapkezelés",
      "titleEn": "Sheet Management",
      "activities": [
        {
          "displayName": "Copy Sheet",
          "inputs": [
            {
              "name": "3. New Sheet Name",
              "required": false,
              "type": "string",
              "default": null,
              "description": "Opcionális. Üres → automatikus névgenerálás"
            }
          ],
          "outputs": [],
          "notes": "Munkalapot másol ugyanazon a fájlon belül."
        },
        {
          "displayName": "Copy Sheet To Workbook",
          "inputs": [
            {
              "name": "1. Source File Path",
              "required": true,
              "type": "string",
              "default": null,
              "description": "Forrás .xlsx fájl"
            },
            {
              "name": "2. Source Sheet Name",
              "required": true,
              "type": "string",
              "default": null,
              "description": "Másolandó lap neve"
            },
            {
              "name": "3. Destination File Path",
              "required": true,
              "type": "string",
              "default": null,
              "description": "Cél .xlsx fájl (ha nem létezik, létrejön)"
            },
            {
              "name": "4. Destination Sheet Name",
              "required": true,
              "type": "string",
              "default": null,
              "description": "Céllapon belüli új lapnév"
            }
          ],
          "outputs": [],
          "notes": "Fájlok közötti másolás. Speciális duális-munkafüzet mintát használ (nem ExcelManager) — nem egymásba ágyazott using blokkokkal."
        },
        {
          "displayName": "Move Sheet To First",
          "inputs": [],
          "outputs": [],
          "notes": "A lapot az 1-es pozícióba helyezi."
        },
        {
          "displayName": "Move Sheet To Last",
          "inputs": [],
          "outputs": [],
          "notes": "A lapot az összes lap számával egyező utolsó pozícióba helyezi."
        },
        {
          "displayName": "Set Sheet Position",
          "inputs": [
            {
              "name": "3. Position",
              "required": true,
              "type": "int",
              "default": null,
              "description": "1-alapú pozícióindex"
            }
          ],
          "outputs": [],
          "notes": "Explicit 1-alapú pozícióindex beállítása a munkafüzeten belül."
        }
      ]
    },
    {
      "id": "format",
      "title": "Formázás",
      "titleEn": "Format",
      "activities": [
        {
          "displayName": "Freeze Panes",
          "inputs": [
            {
              "name": "3. Rows to Freeze",
              "required": false,
              "type": "int",
              "default": "1",
              "description": "Rögzítendő sorok száma (0 = nincs)"
            },
            {
              "name": "4. Columns to Freeze",
              "required": false,
              "type": "int",
              "default": "0",
              "description": "Rögzítendő oszlopok száma (0 = nincs)"
            }
          ],
          "outputs": [],
          "notes": "Mindkét érték lehet 0 (nincs rögzítés abban a tengelyben). A leggyakoribb: Rows=1, Columns=0 (fejlécsor rögzítése)."
        },
        {
          "displayName": "Sort Range",
          "inputs": [
            {
              "name": "3. Range",
              "required": true,
              "type": "string",
              "default": null,
              "description": "Pl. \"A1:D100\""
            },
            {
              "name": "4a. Sort By Column Number",
              "required": false,
              "type": "int?",
              "default": null,
              "description": "Oszlopszám alapján (1-alapú)"
            },
            {
              "name": "4b. Sort By Column Letter",
              "required": false,
              "type": "string",
              "default": null,
              "description": "Oszlopbetű alapján (pl. \"C\")"
            },
            {
              "name": "4c. Sort By Column Name",
              "required": false,
              "type": "string",
              "default": null,
              "description": "Fejléc neve alapján keresés"
            },
            {
              "name": "5. Has Header",
              "required": false,
              "type": "bool",
              "default": "true",
              "description": "A fejlécsor nem kerül a rendezésbe"
            },
            {
              "name": "6. Sort Order",
              "required": false,
              "type": "enum",
              "default": "Ascending",
              "description": "Ascending / Descending"
            }
          ],
          "outputs": [],
          "notes": "A 4a / 4b / 4c közül pontosan egyet kell megadni. A Column Name változat a fejlécsorban keres névazonosság alapján."
        },
        {
          "displayName": "AutoFit Columns",
          "inputs": [
            {
              "name": "3. Columns",
              "required": false,
              "type": "int[]",
              "default": null,
              "description": "Opcionális. Üresen hagyva az összes használt oszlop igazítódik."
            }
          ],
          "outputs": [],
          "notes": "Konkrét oszlopszámok (pl. {1,3,5}) megadásával csak azok igazítódnak."
        },
        {
          "displayName": "Hide Row",
          "inputs": [
            {
              "name": "3. Row Number",
              "required": true,
              "type": "int",
              "default": null,
              "description": "Elrejtendő sor sorszáma (≥ 1)"
            }
          ],
          "outputs": [],
          "notes": "Elrejti a sort. Az adatok megmaradnak, csak vizuálisan rejtett."
        },
        {
          "displayName": "Unhide Row",
          "inputs": [
            {
              "name": "3. Row Number",
              "required": true,
              "type": "int",
              "default": null,
              "description": "Láthatóvá teendő sor sorszáma (≥ 1)"
            }
          ],
          "outputs": [],
          "notes": "Egy rejtett sort láthatóvá tesz."
        },
        {
          "displayName": "Hide Column",
          "inputs": [
            {
              "name": "3. Column Number",
              "required": true,
              "type": "int",
              "default": null,
              "description": "Elrejtendő oszlopszám (≥ 1)"
            }
          ],
          "outputs": [],
          "notes": "Elrejti az oszlopot. Az adatok megmaradnak, csak vizuálisan rejtett."
        },
        {
          "displayName": "Unhide Column",
          "inputs": [
            {
              "name": "3. Column Number",
              "required": true,
              "type": "int",
              "default": null,
              "description": "Láthatóvá teendő oszlopszám (≥ 1)"
            }
          ],
          "outputs": [],
          "notes": "Egy rejtett oszlopot láthatóvá tesz."
        }
      ]
    }
  ],
  "designPrinciples": [
    {
      "title": "Mentési viselkedés",
      "content": "Minden ExcelManager.Dispose() hívás az alábbi beállítással ment. Ez azt jelenti, hogy minden mentéskor az összes képlet újraszámolódik — a fájl mindig aktuális, kiszámított értékeket tartalmaz, amikor Excelben megnyitják.",
      "code": {
        "lang": "csharp",
        "lines": [
          "_workbook.SaveAs(_filePath, new SaveOptions { EvaluateFormulasBeforeSaving = true });"
        ]
      }
    },
    {
      "title": "Bemeneti validáció",
      "content": "Minden bemenet a lehető legkorábban validálódik, típusos kivételekkel és leíró magyar hibaüzenetekkel. A validáció az ExcelManager konstruktorában (fájlútvonal) és az egyes extension metódusokban (null munkalap, érvénytelen sor-/oszlopszám) történik. Soha nem dobunk generikus Exception-t.",
      "code": null
    },
    {
      "title": "Partial class / egy-fájl-egy-metódus konvenció",
      "content": "Minden fájl pontosan egy metódust tartalmaz. A fájlnév megegyezik a metódus nevével. Ez minimalizálja a merge conflict-eket és triviálissá teszi a navigációt.",
      "code": {
        "lang": "text",
        "lines": [
          "MPRT.Excel/",
          "  Excel.ReadRange.cs",
          "  Excel.WriteCell.cs",
          "  Excel.FreezePanes.cs",
          "  ...",
          "  WorksheetExtension.ReadRange.cs",
          "  WorksheetExtension.WriteDataTable.cs",
          "  ..."
        ]
      }
    }
  ]
};
