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
            exportLua_Impl2(data, filePath);
        }

        public static void appendFormatLineEx(StringBuilder sb, int indent, string fmtstr, params object[] args)
        {
            if (indent > 0)
                sb.Append(new String(' ', indent * 4));
            sb.AppendFormat(fmtstr, args);
            sb.AppendLine();
        }

        #region export for simple format
        private static void exportLua_Impl1(TableExcelData data, string filePath)
        {
            var sb = new StringBuilder();
            appendFormatLineEx(sb, 0, "local data = ");
            appendFormatLineEx(sb, 0, "{{");
            foreach (var row in data.Rows)
            {
                appendFormatLineEx(sb, 1, "{{");
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
                    }
                    appendFormatLineEx(sb, 2, "[\"{0}\"] = {1},", hdr.FieldName, s);
                }
                appendFormatLineEx(sb, 1, "}},");
            }
            appendFormatLineEx(sb, 0, "}}");
            sb.AppendLine();
            appendFormatLineEx(sb, 0, "return data");

            File.WriteAllBytes(filePath, Encoding.UTF8.GetBytes(sb.ToString()));
        }
        #endregion

        #region export for utility format
        private static void exportLua_Impl2(TableExcelData data, string filePath)
        {
            var sb = new StringBuilder();

            var ids = new List<string>();
            var keys = new List<string>();

            //comment       ----------------------------------------------------------------------------------
            appendFormatLineEx(sb, 0, "--[[ table define:");
            appendFormatLineEx(sb, 1, "{0,-30} {1,-10} {2}", "name", "type", "desc");
            sb.AppendLine();

            for (int i = 0; i < data.Headers.Count; i++)
            {
                var hdr = data.Headers[i];
                appendFormatLineEx(sb, 1, "{0,-30} {1,-10} {2}", hdr.FieldName, hdr.FieldType, hdr.FieldDesc);
            }
            appendFormatLineEx(sb, 0, "--]]");
            sb.AppendLine();
            sb.AppendLine();

            //comment tips  ----------------------------------------------------------------------------------
            var tipsNum = 0;
            
            for (int i = 0; i < data.Headers.Count; i++)
            {
                var hdr = data.Headers[i];
                if (hdr.FieldDetail.Length > 0)
                {
                    if(tipsNum == 0)
                    {
                        appendFormatLineEx(sb, 0, "--[[ table tips:");
                        sb.AppendLine();
                    }
                    tipsNum++;
                    appendFormatLineEx(sb, 0, "{0}", hdr.FieldName);
                    appendFormatLineEx(sb, 0, "{0,-10}", hdr.FieldDetail);
                    sb.AppendLine();
                }
            }
            if(tipsNum > 0)
            {
                appendFormatLineEx(sb, 0, "--]]");
                sb.AppendLine();
                sb.AppendLine();
            }

            //items         ----------------------------------------------------------------------------------
            appendFormatLineEx(sb, 0, "local items = ");
            appendFormatLineEx(sb, 0, "{{");
            foreach (var row in data.Rows)
            {
                sb.Append("    { ");
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
                    }
                    sb.AppendFormat("{0} = {1}, ", hdr.FieldName, s);

                    if (hdr.FieldName == "id")
                        ids.Add(s);
                    else if (hdr.FieldName == "key")
                        keys.Add(s);
                }
                sb.AppendLine("},");
            }
            appendFormatLineEx(sb, 0, "}}");
            sb.AppendLine();

            //idItems       ----------------------------------------------------------------------------------
            appendFormatLineEx(sb, 0, "local idItems = ");
            appendFormatLineEx(sb, 0, "{{");
            for (int i = 0; i < ids.Count; i++)
            {
                appendFormatLineEx(sb, 1, "[{0}] = items[{1}],", ids[i], i + 1);
            }
            appendFormatLineEx(sb, 0, "}}");
            sb.AppendLine();

            //keyItems      ----------------------------------------------------------------------------------
            appendFormatLineEx(sb, 0, "local keyItems = ");
            appendFormatLineEx(sb, 0, "{{");
            for (int i = 0; i < keys.Count; i++)
            {
                appendFormatLineEx(sb, 1, "[{0}] = items[{1}],", keys[i], i + 1);
            }
            appendFormatLineEx(sb, 0, "}}");
            sb.AppendLine();

            sb.Append( BuildFunction(Path.GetFileName(filePath))) ;

            File.WriteAllBytes(filePath, Encoding.UTF8.GetBytes(sb.ToString()));
        }
        #endregion

        private static string BuildFunction( string fileName )
        {
            var sb = new StringBuilder();
            sb.Append(
@"
local data = { Items = items, IdItems = idItems, KeyItems = keyItems, }
");
            sb.AppendFormat(
@"
function data:GetById( id, prop )
    local dat = self.IdItems[id];
    if dat == nil then
        sGlobal:Print( ""{0} GetById error, invalid id: ""..id );
        return id;
    end
    if prop == nil then
        return dat;
    end
    if dat[prop] == nil then
        sGlobal:Print( ""{0} GetById error, invalid prop: ""..prop );
        return dat;
    end
    return dat[prop];
end

function data:GetByKey( key, prop )
    local dat = self.KeyItems[key];
    if dat == nil then
        sGlobal:Print( ""{0} GetByKey error, invalid key: ""..key );
        return id;
    end
    if prop == nil then
        return dat;
    end
    if dat[prop] == nil then
        sGlobal:Print( ""{0} GetByKey error, invalid prop: ""..prop );
        return dat;
    end
    return dat[prop];
end

function data:GetCount()
    return #self.IdItems;
end

return data
", Path.GetFileName(fileName));
            return sb.ToString();
        }
    }
}

//
