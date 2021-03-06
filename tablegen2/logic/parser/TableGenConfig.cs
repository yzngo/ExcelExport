﻿using System.Collections.Generic;
namespace tablegen2.logic
{
    public class TableGenConfig
    {
        public string ExcelDir = string.Empty;
        public TableExportFormat ExportFormat = TableExportFormat.Unknown;
        public string ExportDir = string.Empty;

        public string SheetNameForField = string.Empty;
        public string SheetNameForData = string.Empty;

        public bool OutputLuaWithIndent = true;
        public bool FitUnity3D = false;

        public Dictionary<string, int> ExcelDirHistory;
        public Dictionary<string, int> ExportDirHistory;
    }
}
