using System.Collections.Generic;
using System.Linq;

namespace tablegen2.logic
{
    public class TableExcelData
    {
        private List<TableExcelHeader> headers_ = new List<TableExcelHeader>();
        private List<TableExcelRow> rows_ = new List<TableExcelRow>();

        public TableExcelData()
        {
        }

        public TableExcelData(IEnumerable<TableExcelHeader> headers, IEnumerable<TableExcelRow> rows)
        {
            headers_ = headers.ToList();
            rows_ = rows.ToList();
        }

        public List<TableExcelHeader> Headers
        {
            get { return headers_; }
        }

        public List<TableExcelRow> Rows
        {
            get { return rows_; }
        }

        public bool checkUnique(out string errmsg)
        {
            int idx1 = headers_.FindIndex(a => a.FieldName.Equals("id"));
            int idx2 = headers_.FindIndex(a => a.FieldName.Equals("key"));
            var ids = new HashSet<int>();
            var keys = new HashSet<string>();
            for (int i = 0; i < rows_.Count; i++)
            {
                var row = rows_[i];
                var strId = row.StrList[idx1];
                var strKeyName = row.StrList[idx2];

                int id;
                if (!int.TryParse(strId, out id))
                {
                    errmsg = string.Format("第{0}行id值非法，须为数字类型：{1}", i + 2, strId);
                    return false;
                }

                if (string.IsNullOrEmpty(strKeyName))
                {
                    errmsg = string.Format("第{0}行key值为空", i + 2);
                    return false;
                }

                if (ids.Contains(id))
                {
                    errmsg = string.Format("第{0}行id值已存在：{1}", i + 2, strId);
                    return false;
                }

                if (keys.Contains(strKeyName))
                {
                    errmsg = string.Format("第{0}行key值已存在：{1}", i + 2, strKeyName);
                    return false;
                }

                ids.Add(id);
                keys.Add(strKeyName);
            }

            errmsg = string.Empty;
            return true;
        }

        public bool checkDataType(out string errmsg)
        {
            var nullValue = 0;
            var invalidValue = 0;
            for(int i = 0; i < rows_.Count; ++i )
            {
                for (int j = 0; j < headers_.Count; j++)
                {
                    var hdr = headers_[j];
                    var val = rows_[i].StrList[j];
                    
                    string s = string.Empty;
                    if(string.IsNullOrEmpty(val) && hdr.FieldType != "group")
                    {
                        nullValue += 1;
                    }
                    else
                    {
                        switch (hdr.FieldType)
                        {
                            case "string":
                                break;
                            case "int":
                                {
                                    if (!int.TryParse(val, out _))
                                    {
                                        invalidValue += 1;
                                        Log.Err(string.Format("第 {0} 行 第 {1} 列 数据 {2} 类型不匹配，应为 int 型", i + 2, j + 1,  val ));
                                    }
                                }
                                break;
                            case "double":
                                {
                                    if (!double.TryParse(val, out _))
                                    {
                                        invalidValue += 1;
                                        Log.Err(string.Format("第 {0} 行 第 {1} 列 数据 {2} 类型不匹配，应为 double 型", i + 2, j + 1, val ));
                                    }
                                }
                                break;
                            case "bool":
                                {
                                    if (!bool.TryParse(val, out _))
                                    {
                                        invalidValue += 1;
                                        Log.Err(string.Format("第 {0} 行 第 {1} 列 数据 {2} 类型不匹配，应为 bool 型", i + 2, j + 1, val));
                                    }
                                }
                                break;
                            case "group":
                                {
                                    s = string.Format("{{{0}}}", val);
                                }
                                break;
                            case "color":
                                {
                                    var substring = val.Substring(0, 2);
                                    if (val.Length != 8 || string.Compare(substring, "0x") != 0)
                                    {

                                        invalidValue += 1;
                                        Log.Err(string.Format("第 {0} 行 第 {1} 列 数据 {2} 类型不匹配，应为 color 型", i + 2, j + 1, val));
                                    }
                                    else
                                    {
                                        IList<char> HexSet = new List<char>() { '0','1','2','3','4','5','6','7','8','9','A','B','C','D','E','F','a','b','c','d','e','f' };
                                        for (int counter = 2; counter < val.Length; ++counter)
                                        {
                                            if (!HexSet.Contains<char>(val[counter]))
                                            {
                                                invalidValue += 1;
                                                Log.Err(string.Format("第 {0} 行 第 {1} 列 数据 {2} 类型不匹配，应为 color 型", i + 2, j + 1, val));
                                                break;
                                            }
                                        }                          
                                    }
                                }
                                break;
                        }
                    }

                }
            }

            if (nullValue > 0)
            {
                Log.Wrn( string.Format("此表格中存在 {0} 个空值", nullValue) );
            }

            if (invalidValue > 0)
            {
                errmsg = string.Format("此表格中一共有 {0} 个类型错误的字段", invalidValue);
                return false;
            }
            else
            {
                errmsg = string.Empty;
                return true;
            }
        }
    }
}
