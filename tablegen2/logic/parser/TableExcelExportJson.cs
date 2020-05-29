using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using System.IO;




namespace tablegen2.logic
{
    public static class TableExcelExportJson
    {
        public class CustomJsonTextWriter : JsonTextWriter
        {
            public CustomJsonTextWriter(TextWriter writer) : base(writer)
            {
            }

            protected override void WriteIndent()
            {
                if (WriteState != WriteState.Array)
                    base.WriteIndent();
                else
                    WriteIndentSpace();
            }
        }

        public static string SerializeWithCustomIndenting(object obj)
        {
            using (StringWriter sw = new StringWriter())
            using (JsonWriter jw = new CustomJsonTextWriter(sw))
            {
                jw.Formatting = Formatting.Indented;
                JsonSerializer ser = new JsonSerializer();
                ser.Serialize(jw, obj);
                return sw.ToString();
            }
        }

        public static void exportExcelFile(TableExcelData data, string filePath)
        {
            List<Dictionary<string, object>> lst = ExportData(data);

            // Dictionary<string, List<Dictionary<string, object>>> temp = new Dictionary<string, List<Dictionary<string, object>>>();
            // var fileName = Path.GetFileNameWithoutExtension(filePath);
            // temp[fileName] = lst;
            var indent = AppData.Config.OutputLuaWithIndent ? Formatting.Indented : Formatting.None;
            string output = SerializeWithCustomIndenting(lst);
            //string output = JsonConvert.SerializeObject(lst, indent);
            File.WriteAllBytes(filePath, Encoding.UTF8.GetBytes(output));
        }

        public static List<Dictionary<string, object>> ExportData(TableExcelData data)
        {
            List<Dictionary<string, object>> lst = new List<Dictionary<string, object>>();

            foreach (var row in data.Rows)
            {
                var r = new Dictionary<string, object>();

                for (int i = 0; i < data.Headers.Count; i++)
                {
                    var hdr = data.Headers[i];
                    var val = row.StrList[i];

                    if (string.IsNullOrEmpty(val) && !(hdr.FieldType == "group" || hdr.FieldType == "string" || hdr.FieldType == "table"))
                    {
                        continue;
                    }
                    object obj = null;
                    switch (hdr.FieldType)
                    {
                        case "string":
                        case "string(nil)":
                        case "color":
                            obj = val;
                            break;
                        case "int":
                            {
                                int.TryParse(val, out int n);
                                obj = n;
                            }
                            break;
                        case "double":
                            {
                                double.TryParse(val, out double n);
                                obj = n;
                            }
                            break;
                        case "bool":
                            {
                                bool.TryParse(val, out bool n);
                                obj = n;
                            }
                            break;
                        case "group":
                            {
                                var str = val.Split(',');
                                bool succeed = false;
                                {
                                    var numlist = new List<double>();
                                    foreach (var s in str)
                                    {
                                        succeed = double.TryParse(s, out double n);
                                        numlist.Add(n);
                                    }
                                    obj = numlist;
                                }
                                
                                if (succeed == false)
                                {
                                    var strlist = new List<string>();
                                    foreach (var s in str)
                                    {
                                        strlist.Add(s.Replace("\"", string.Empty));
                                    }
                                    obj = strlist;
                                }
                            }
                            break;
                        case "table":
                            {
                                obj = ExportDataSub(data.ChildData[hdr.FieldName], row.StrList[1]);
                                break;
                            }
                    }
                    if (hdr.FieldName == "id")
                    {
                        r["index"] = obj;
                    } else if (hdr.FieldName == "key")
                    {
                        r["id"] = obj;
                    } else
                    {
                        r[hdr.FieldName] = obj;
                    }
                }
                lst.Add(r);
            }
            return lst;
        }

        public static Dictionary<string, object> ExportDataSub(TableExcelData data, string key)
        {
            var r = new Dictionary<string, object>();
            foreach (var row in data.Rows)
            {
                if (key != string.Empty && key != row.StrList[1])
                {
                    continue;
                }
                for (int i = 0; i < data.Headers.Count; i++)
                {
                    var hdr = data.Headers[i];
                    var val = row.StrList[i];
                    if (key != string.Empty && (hdr.FieldName == "id" || hdr.FieldName == "key"))
                    {
                        continue;
                    }
                    if (string.IsNullOrEmpty(val) && !(hdr.FieldType == "group" || hdr.FieldType == "string" || hdr.FieldType == "table"))
                    {
                        continue;
                    }
                    object obj = null;
                    switch (hdr.FieldType)
                    {
                        case "string":
                        case "string(nil)":
                        case "color":
                            obj = val;
                            break;
                        case "int":
                            {
                                int.TryParse(val, out int n);
                                obj = n;
                            }
                            break;
                        case "double":
                            {
                                double.TryParse(val, out double n);
                                obj = n;
                            }
                            break;
                        case "bool":
                            {
                                bool.TryParse(val, out bool n);
                                obj = n;
                            }
                            break;
                        case "group":
                            {
                                var str = val.Split(',');
                                bool succeed = false;
                                {
                                    var numlist = new List<double>();
                                    foreach (var s in str)
                                    {
                                        succeed = double.TryParse(s, out double n);
                                        numlist.Add(n);
                                    }
                                    obj = numlist;
                                }

                                if (succeed == false)
                                {
                                    var strlist = new List<string>();
                                    foreach (var s in str)
                                    {
                                        strlist.Add(s.Replace("\"", string.Empty));
                                    }
                                    obj = strlist;
                                }
                            }
                            break;
                        case "table":
                            {
                                obj = ExportDataSub(data.ChildData[hdr.FieldName], row.StrList[1]);
                                break;
                            }
                    }
                    r[hdr.FieldName] = obj;
                }
            }
            return r;
        }
    }
}
