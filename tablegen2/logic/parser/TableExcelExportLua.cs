using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace tablegen2.logic
{
    public static class TableExcelExportLua
    {
        public static void exportExcelFile(TableExcelData data, string filePath)
        {
            exportLua_Impl(data, filePath);
        }

        public static void appendFormatLineEx(StringBuilder sb, int indent, string fmtstr, params object[] args)
        {
            if (indent > 0)
                sb.Append(new String(' ', indent * 4));
            sb.AppendFormat(fmtstr, args);
            sb.AppendLine();
        }

        private static string BuildCommentString( List<TableExcelHeader> headers )
        {
            var luaString = new StringBuilder();
            appendFormatLineEx(luaString, 0, "--[[ table define:");
            appendFormatLineEx(luaString, 1, "{0,-30} {1,-10} {2}", "name", "type", "desc");
            luaString.AppendLine();

            for (int i = 0; i < headers.Count; i++)
            {
                var hdr = headers[i];
                appendFormatLineEx(luaString, 1, "{0,-30} {1,-10} {2}", hdr.FieldName, hdr.FieldType, hdr.FieldDesc);
            }
            appendFormatLineEx(luaString, 0, "--]]");
            luaString.AppendLine();
            luaString.AppendLine();

            var tipsNum = 0;

            for (int i = 0; i < headers.Count; i++)
            {
                var hdr = headers[i];
                if (hdr.FieldDetail.Length > 0)
                {
                    if (tipsNum == 0)
                    {
                        appendFormatLineEx(luaString, 0, "--[[ table tips:");
                        luaString.AppendLine();
                    }
                    tipsNum++;
                    appendFormatLineEx(luaString, 0, "{0}", hdr.FieldName);
                    appendFormatLineEx(luaString, 0, "{0,-10}", hdr.FieldDetail);
                    luaString.AppendLine();
                }
            }
            if (tipsNum > 0)
            {
                appendFormatLineEx(luaString, 0, "--]]");
                luaString.AppendLine();
                luaString.AppendLine();
            }
            return luaString.ToString();
        }

        private static string BuildDataString(TableExcelData data, string key)
        {
            var luaString = new StringBuilder();

            foreach (var row in data.Rows)
            {
                if (key == string.Empty)
                {
                    luaString.Append("    ");
                }
                if (key != string.Empty && key != row.StrList[1])
                    continue;
                luaString.Append("{ ");
                for (int i = 0; i < data.Headers.Count; i++)
                {
                    var hdr = data.Headers[i];
                    var val = row.StrList[i];

                    string s = string.Empty;
                    switch (hdr.FieldType)
                    {
                        case "string":
                            s = string.Format("\"{0}\"", val);
                            break;
                        case "int":
                            {
                                int n = 0;
                                int.TryParse(val, out n);
                                s = n.ToString();
                            }
                            break;
                        case "double":
                            {
                                double n = 0;
                                double.TryParse(val, out n);
                                s = n.ToString();
                            }
                            break;
                        case "bool":
                            {
                                bool n = true;
                                bool.TryParse(val, out n);
                                s = n == true ? "true" : "false";
                            }
                            break;
                        case "group":
                            {
                                s = string.Format("{{{0}}}", val);
                            }
                            break;
                        case "color":
                            {
                                s = string.Format("{0:X}", val);
                            }
                            break;
                        case "table":
                            {
                                s = BuildDataString(data.ChildData[hdr.FieldName], row.StrList[1]);
                            }
                            break;
                    }
                    if (key != string.Empty && (hdr.FieldName == "id" || hdr.FieldName == "key"))
                        continue;

                    luaString.AppendFormat("{0} = {1}, ", hdr.FieldName, s);
                    
                }
                if(key == string.Empty)
                {
                    luaString.AppendLine("},");
                }
                else
                {
                    luaString.Append("}");
                }
                

            }
            return luaString.ToString();
        }

        #region export lua table
        private static void exportLua_Impl(TableExcelData data, string filePath)
        {
            var luaString = new StringBuilder();
            luaString.Append(BuildCommentString(data.Headers));

            appendFormatLineEx(luaString, 0, "local items = ");
            appendFormatLineEx(luaString, 0, "{{");

            luaString.Append(BuildDataString(data, string.Empty));

            appendFormatLineEx(luaString, 0, "}}");
            luaString.AppendLine();
            luaString.Append(BuildItemString(data));
            luaString.Append( BuildFuncString(Path.GetFileName(filePath))) ;
            File.WriteAllBytes(filePath, Encoding.UTF8.GetBytes(luaString.ToString()));
        }
        #endregion
        private static string BuildItemString(TableExcelData data)
        {
            var ids = new List<string>();
            var keys = new List<string>();
            var luaString = new StringBuilder();
            foreach (var row in data.Rows)
            {
                for (int i = 0; i < data.Headers.Count; i++)
                {
                    var hdr = data.Headers[i];
                    var val = row.StrList[i];
                    string s = string.Empty;
                    if (hdr.FieldName == "id")
                    {
                        int n = 0;
                        int.TryParse(val, out n);
                        s = n.ToString();
                        ids.Add(s);
                    }
                    else if (hdr.FieldName == "key")
                    {
                        s = string.Format("\"{0}\"", val);
                        keys.Add(s);
                    }

                }
            }

            appendFormatLineEx(luaString, 0, "local idItems = ");
            appendFormatLineEx(luaString, 0, "{{");
            for (int i = 0; i < ids.Count; i++)
            {
                appendFormatLineEx(luaString, 1, "[{0}] = items[{1}],", ids[i], i + 1);
            }
            appendFormatLineEx(luaString, 0, "}}");
            luaString.AppendLine();

            appendFormatLineEx(luaString, 0, "local keyItems = ");
            appendFormatLineEx(luaString, 0, "{{");
            for (int i = 0; i < keys.Count; i++)
            {
                appendFormatLineEx(luaString, 1, "[{0}] = items[{1}],", keys[i], i + 1);
            }
            appendFormatLineEx(luaString, 0, "}}");
            luaString.AppendLine();
            return luaString.ToString();
        }

        private static string BuildFuncString( string fileName )
        {
            var luaString = new StringBuilder();
            luaString.Append(
@"
local data = { Items = items, IdItems = idItems, KeyItems = keyItems, }
");
            luaString.AppendFormat(
@"
function data:GetById( id, prop )
    local item = self.IdItems[id];
    if item == nil then
        sGlobal:Print( ""{0} GetById nil item: ""..id );
        return id;
    end
    if prop == nil then
        return item;
    end
    if item[prop] == nil then
        sGlobal:Print( ""{0} GetById nil prop: ""..prop );
        return item;
    end
    return item[prop];
end

function data:GetByKey( key, prop )
    local item = self.KeyItems[key];
    if item == nil then
        sGlobal:Print( ""{0} GetByKey nil key: ""..key );
        return key;
    end
    if prop == nil then
        return item;
    end
    if item[prop] == nil then
        sGlobal:Print( ""{0} GetByKey nil prop: ""..prop );
        return item;
    end
    return item[prop];
end

function data:GetCount()
    return #self.IdItems;
end

return data
", Path.GetFileName(fileName));
            return luaString.ToString();
        }
    }
}
