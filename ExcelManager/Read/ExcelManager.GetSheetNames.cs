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
        /// Visszaadja a munkafüzetben található összes munkalap nevét.
        /// </summary>
        /// <returns>A munkalapnevek listája (IEnumerable<string>).</returns>
        public IEnumerable<string> GetSheetNames()
        {
            // A logikát a WorkbookExtensions-ben definiált bővítő metódusra delegálja.
            return _workbook.GetSheetNames();
        }

    }
}
