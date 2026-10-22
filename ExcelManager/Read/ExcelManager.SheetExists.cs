using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;

namespace MPRT.Excel
{
    /// <summary>
    /// Egy Excel munkafüzetet kezelő osztály, amely biztosítja a robusztus fájlkezelést és
    /// lehetővé teszi a láncolt (fluent) műveleteket.
    /// Implementálja az IDisposable interfészt a biztonságos erőforrás-kezelés érdekében.
    /// Használata 'using' blokkban javasolt.
    /// </summary>
    public partial class ExcelManager : IDisposable
    {
        /// <summary>
        /// Ellenőrzi, hogy egy adott nevű munkalap létezik-e a munkafüzetben.
        /// </summary>
        /// <param name="sheetName">A keresendő munkalap neve.</param>
        /// <returns>Igaz (true), ha a munkalap létezik; egyébként hamis (false).</returns>
        public bool SheetExists(string sheetName)
        {
            // A logikát a WorkbookExtensions-ben definiált bővítő metódusra delegálja.
            // Ez a tiszta felelősségi köröket (separation of concerns) valósítja meg.
            return _workbook.SheetExists(sheetName);
        }
    }
}
