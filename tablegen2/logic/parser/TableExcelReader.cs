using System;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using NPOI.XSSF.UserModel;

namespace tablegen2.logic
{
    public static class TableExcelReader
    {
        public static TableExcelData loadFromExcel(string filePath)
        {
            if (!File.Exists(filePath))
                throw new Exception(string.Format("{0} 文件不存在！", filePath));

            var ext = Path.GetExtension(filePath).ToLower();
            if (ext != ".xls" && ext != ".xlsx")
                throw new Exception(string.Format("无法识别的文件扩展名 {0}", ext));

            var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var workbook = ext == ".xls" ? (IWorkbook)new HSSFWorkbook(fs) : (IWorkbook)new XSSFWorkbook(fs);
            fs.Close();

            var def = AppData.Config.SheetNameForField;
            var data = AppData.Config.SheetNameForData;
            return _readDataFromWorkbook(workbook, def, 1, 0, data );
        }

        private static TableExcelData _readDataFromWorkbook(IWorkbook wb, string defSheetName, int row, int column, string dataSheetName )
        {
            var rows = new List<TableExcelRow>();

            var sheet1 = wb.GetSheet(defSheetName);
            if (sheet1 == null)
                throw new Exception(string.Format("'{0}'工作簿不存在", defSheetName));

            var sheet2 = wb.GetSheet(dataSheetName);
            if (sheet2 == null)
                throw new Exception(string.Format("'{0}'工作簿不存在", dataSheetName));

            //加载字段
           var headers = _readHeadersFromDefSheet(sheet1, row, column, defSheetName);
            var h1 = headers.Find(a => a.FieldName.ToLower() == "id");
            if (h1 == null)
                throw new Exception(string.Format("'{0}'工作簿{1}行{2}列缺失id字段！", defSheetName, row, column + 1));

            var h2 = headers.Find(a => a.FieldName.ToLower() == "key");
            if (h2 == null)
                throw new Exception(string.Format("'{0}'工作簿{1}行{2}列缺失key字段！", defSheetName, row, column + 1));

            //加载数据
            var headers2 = _readHeadersFromDataSheet(sheet2, dataSheetName);
            var headerIndexes = new int[headers.Count];

            _checkFieldsSame(headers, headers2, headerIndexes, dataSheetName);

            foreach (var ds in _readDataFromDataSheet(sheet2, headers2.Count))
            {
                var rowData = new List<string>();
                for (int i = 0; i < headers.Count; i++)
                {
                    var idx = headerIndexes[i];
                    rowData.Add(ds[idx]);
                }
                rows.Add(new TableExcelRow() { StrList = rowData });
            }

            var data = new TableExcelData(headers, rows);

            for (int r = row; r <= sheet1.LastRowNum; r++)
            {
                var rd = sheet1.GetRow(r);
                if (rd == null)
                    continue;
                var str1 = _convertCellToString(rd.GetCell(column + 0));
                var str2 = _convertCellToString(rd.GetCell(column + 1));
                if (string.IsNullOrEmpty(str2))
                    continue;
                else
                {
                    var cell = rd.GetCell(column + 1);
                    if (cell.Hyperlink != null && cell.Hyperlink.Type == HyperlinkType.Document)
                    {
                        string s = cell.Hyperlink.Address;
                        char[] addr = s.ToCharArray();
                        int x = 0;
                        int y = 0;
                        if(s.Length == 6)
                        {
                            x = addr[5] - '0';
                            y = addr[4] - 'A';
                        }
                        else
                        {
                            x = addr[6] - '0';
                            y = (addr[4] - 'A' + 1) * 26 + addr[5] - 'A';
                        }
                        data.ChildData.Add(str1, _readDataFromWorkbook(wb, defSheetName, x, y, str1));
                    }
                }
            }  
            return data;
        }

        private static string _convertCellToString(ICell cell)
        {
            
            string r = string.Empty;
            if (cell != null)
            {
                switch (cell.CellType)
                {
                    case CellType.Boolean:
                        r = cell.BooleanCellValue.ToString();
                        break;
                    case CellType.Numeric:
                        r = cell.NumericCellValue.ToString();
                        break;
                    case CellType.Formula:
                        r = cell.CellFormula;
                        break;
                    default:
                        r = cell.StringCellValue;
                        break;
                }
            }
            return r;
        }

        private static List<TableExcelHeader> _readHeadersFromDefSheet( ISheet sheet, int row, int column, string sheetName )
        {
            
            var headers = new List<TableExcelHeader>();
            for (int r = row; r <= sheet.LastRowNum; r++)
            {
                var rd = sheet.GetRow(r);
                if (rd == null)
                    continue;
                var str1 = _convertCellToString(rd.GetCell(column + 0));
                var str2 = _convertCellToString(rd.GetCell(column + 1));
                var str3 = _convertCellToString(rd.GetCell(column + 2));
                var str4 = _convertCellToString(rd.GetCell(column + 3));
                
                if (string.IsNullOrEmpty(str1) && string.IsNullOrEmpty(str2) && string.IsNullOrEmpty(str3) )
                    continue;

                if (!string.IsNullOrEmpty(str1 ) && !string.IsNullOrEmpty(str2))
                {
                    headers.Add(new TableExcelHeader()
                    {
                        FieldName = str1,
                        FieldType = str2,
                        FieldDesc = str3,
                        FieldDetail = str4,
                    });
                    continue;
                }

                throw new Exception(string.Format(
                    "'{0}'工作簿中第{1}行第{2}列数据异常，有缺失！", sheetName, r + 1, column + 1));
            }
            return headers;
        }
        
        private static List<string> _readHeadersFromDataSheet(ISheet sheet, string sheetName)
        {
            var r = new List<string>();
            var rd = sheet.GetRow(0);
            for (int i = 0; i < rd.LastCellNum; i++)
            {
                var cell = rd.GetCell(i);
                r.Add(cell != null ? cell.StringCellValue : string.Empty);
            }
            for (int i = r.Count - 1; i >= 0; i--)
            {
                if (string.IsNullOrEmpty(r[i]))
                    r.RemoveAt(i);
                else
                    break;
            }
            int idx = r.IndexOf(string.Empty);
            if (idx >= 0)
                throw new Exception(string.Format(
                    "'{0}'工作簿中第1行第{1}列字段名称非法", sheetName, idx + 1));
            return r;
        }

        private static void _checkFieldsSame(List<TableExcelHeader> headers1, List<string> headers2, int[] indexes, string sheetName)
        {
            for (int i = 0; i < headers1.Count; i++)
            {
                var hd = headers1[i];
                var idx = headers2.IndexOf(hd.FieldName);
                if (idx < 0)
                    throw new Exception(string.Format("'{0}'工作簿中不存在字段'{1}'所对应的列", sheetName, hd.FieldName));
                indexes[i] = idx;
            }
            if (headers1.Count < headers2.Count)
            {
                foreach (var s in headers2)
                {
                    if (headers1.Find(a => a.FieldName == s) == null)
                        throw new Exception(string.Format("'{0}'工作簿中包含多余的数据列'{1}'", sheetName, s));
                }
            }
        }

        private static IEnumerable<List<string>> _readDataFromDataSheet(ISheet sheet, int columns)
        {
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                var rd = sheet.GetRow(i);
                if (rd == null)
                    continue;

                var ds = new List<string>();
                bool is_all_empty = true;
                for (int c = 0; c < columns; c++)
                {
                    var val = _convertCellToString(rd.GetCell(c));
                    if (!string.IsNullOrEmpty(val))
                        is_all_empty = false;
                    ds.Add(val);
                }

                if (!is_all_empty)
                    yield return ds;
            }
        }
    }
}
