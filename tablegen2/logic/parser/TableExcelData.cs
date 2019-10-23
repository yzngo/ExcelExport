using System.Collections.Generic;
using System.Linq;

namespace tablegen2.logic
{
    public class TableExcelData
    {
        public TableExcelData()
        {
        }

        public TableExcelData(IEnumerable<TableExcelHeader> headers, IEnumerable<TableExcelRow> rows)
        {
            Headers = headers.ToList();
            Rows = rows.ToList();
        }

        public List<TableExcelHeader> Headers { get; } = new List<TableExcelHeader>();
        public List<TableExcelRow> Rows { get; } = new List<TableExcelRow>();
        public Dictionary<string, TableExcelData> ChildData { get; set; } = new Dictionary<string, TableExcelData>();

        public bool checkUnique(out string errmsg)
        {
            int idx1 = Headers.FindIndex(a => a.FieldName.Equals("id"));
            int idx2 = Headers.FindIndex(a => a.FieldName.Equals("key"));
            var ids = new HashSet<int>();
            var keys = new HashSet<string>();
            for (int i = 0; i < Rows.Count; i++)
            {
                var row = Rows[i];
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
            //for (int i = 0; i < Headers.Count; i++)
            //{
            //    var hdr = Headers[i];
            //    if( hdr.FieldType == "table" )
            //    {
            //        ChildData[i].checkUnique(out errmsg);
            //        if(errmsg != string.Empty)
            //        {
            //            return false;
            //        }
            //    } 
            //}

            errmsg = string.Empty;
            return true;
        }

        public bool checkDataType(out string errmsg)
        {
            var nullValue = 0;
            var invalidValue = 0;
            for(int i = 0; i < Headers.Count; ++i)
            {
                var hdr = Headers[i];
                IList<string> field = new List<string>() { "string","int", "double", "bool", "color", "group", "table", };
                if (!field.Contains<string>(hdr.FieldType))
                {
                    invalidValue++;
                    Log.Err(string.Format("第 {0} 列 数据类型错误", i + 1));
                }
            }

            for(int i = 0; i < Rows.Count; ++i)
            {
                for (int j = 0; j < Headers.Count; ++j)
                {
                    var hdr = Headers[j];
                    var val = Rows[i].StrList[j];
                    
                    if(string.IsNullOrEmpty(val) && hdr.FieldType != "group")
                    {
                        nullValue++;
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
                                        invalidValue ++;
                                        Log.Err(string.Format("第 {0} 行 第 {1} 列 数据 {2} 类型不匹配，应为 int 型", i + 2, j + 1,  val ));
                                    }
                                }
                                break;
                            case "double":
                                {
                                    if (!double.TryParse(val, out _))
                                    {
                                        invalidValue++;
                                        Log.Err(string.Format("第 {0} 行 第 {1} 列 数据 {2} 类型不匹配，应为 double 型", i + 2, j + 1, val ));
                                    }
                                }
                                break;
                            case "bool":
                                {
                                    if (!bool.TryParse(val, out _))
                                    {
                                        invalidValue++;
                                        Log.Err(string.Format("第 {0} 行 第 {1} 列 数据 {2} 类型不匹配，应为 bool 型", i + 2, j + 1, val));
                                    }
                                }
                                break;
                            case "color":
                                {
                                    var substring = val.Substring(0, 2);
                                    if (val.Length != 8 || string.Compare(substring, "0x") != 0)
                                    {

                                        invalidValue++;
                                        Log.Err(string.Format("第 {0} 行 第 {1} 列 数据 {2} 类型不匹配，应为 color 型", i + 2, j + 1, val));
                                    }
                                    else
                                    {
                                        IList<char> HexSet = new List<char>() { '0','1','2','3','4','5','6','7','8','9','A','B','C','D','E','F','a','b','c','d','e','f' };
                                        for (int counter = 2; counter < val.Length; ++counter)
                                        {
                                            if (!HexSet.Contains<char>(val[counter]))
                                            {
                                                invalidValue++;
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
                errmsg = string.Format("此表格中一共有 {0} 个类型错误", invalidValue);
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
