using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using System.IO;

namespace tablegen2.logic
{
    public static class TableExcelExportJson
    {
        public static void exportExcelFile(TableExcelData data, string filePath)
        {
            List<Dictionary<string, object>> lst = data.Rows.Select(a =>
            {
                var r = new Dictionary<string, object>();
                for (int i = 0; i < data.Headers.Count; i++)
                {
                    var hdr = data.Headers[i];
                    var val = a.StrList[i];
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
                                var num = new List<double>();
                                foreach (var s in str)
                                {
                                    double.TryParse(s, out double n);
                                    num.Add(n);
                                }
                                obj = num;
                            }
                            break;
                        case "table":
                            {
                                break;
                            }
                    }
                    r[hdr.FieldName] = obj;
                }
                return r;
            }).ToList();

            string output = JsonConvert.SerializeObject(lst, Formatting.Indented);
            File.WriteAllBytes(filePath, Encoding.UTF8.GetBytes(output));
        }
    }
}
