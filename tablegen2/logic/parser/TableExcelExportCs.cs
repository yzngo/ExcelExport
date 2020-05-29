﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace tablegen2.logic
{
    public static class TableExcelExportCs
    {
        public static void exportExcelFile(TableExcelData data, string filePath)
        {
            exportCS_Impl(data, filePath);
        }

        #region export cs class
        private static void exportCS_Impl(TableExcelData data, string filePath)
        {
            var csString = new StringBuilder();
            csString.Append(BuildPrologue(data.Headers));
            // csString.Append(BuildCommentString(data.Headers));

           // appendFormatLineEx(csString, 0, "local items =");
           // appendFormatLineEx(csString, 0, "{{");

            //csString.Append(BuildDataString(data, string.Empty, 2));

            //appendFormatLineEx(csString, 0, "}}");
            //csString.AppendLine();
            //csString.Append(BuildItemString(data));
            csString.Append(BuildFuncString(Path.GetFileName(filePath)));
            File.WriteAllBytes(filePath, Encoding.UTF8.GetBytes(csString.ToString()));
        }
        #endregion

        public static void appendFormatLineEx(this StringBuilder sb, int indent, string fmtstr, params object[] args)
        {
            if (indent > 0)
                sb.Append(new String(' ', indent * 4));
            sb.AppendFormat(fmtstr, args);
            sb.AppendLine();
        }

        private static string BuildPrologue( List<TableExcelHeader> headers )
        {
            var prologue = new StringBuilder();
            prologue.Append(
@"
namespace Feamber.Data
{
    using System;
    using System.Collections.Generic;

    using System.Globalization;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Converters;
}
");
            return prologue.ToString();
        }

        private static string BuildCommentString( List<TableExcelHeader> headers )
        {
            var luaString = new StringBuilder();
            luaString.Append(
@"--[[
item define:
");
            luaString.appendFormatLineEx( 1, "{0,-25} {1,-15} {2}", "____name", "____type", "____desc");

            for (int i = 0; i < headers.Count; i++)
            {
                var hdr = headers[i];

                var name = hdr.FieldName;
                var desc = hdr.FieldDesc;
                if (hdr.FieldName == "id")
                {
                    name = "index";
                    desc = "唯一数字索引";
                }
                else if (hdr.FieldName == "key")
                {
                    name = "id";
                    desc = "唯一字符串索引";
                }
                    
                appendFormatLineEx(luaString, 1, "{0,-25} {1,-15} {2}", name, hdr.FieldType, desc);
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
            appendFormatLineEx(luaString, 0, "--To consider the programmer's habits,replace the 'id' and 'key' in Excel with 'index' and 'id'");
            return luaString.ToString();
        }

        private static string BuildDataString(TableExcelData data, string key, int iDeep)
        {
            bool bWithIndent = AppData.Config.OutputLuaWithIndent;
            var csString = new StringBuilder();
            int lbracket = 0, rbracket = 0; ;
            foreach (var row in data.Rows)
            {
                if (key == string.Empty)
                {
                    if(bWithIndent == true)
                    {
                        for (int t = 1; t <= iDeep - 1; t++)
                        {
                            csString.Append("    ");
                        }
                    }
                    else
                    {
                        csString.Append("    ");
                    }
 
                }
                if (key != string.Empty && key != row.StrList[1])
                    continue;

                if(bWithIndent == true)
                {
                    csString.Append("{\n");
                }
                else
                {
                    csString.Append("{");
                }
                lbracket++;
                for (int i = 0; i < data.Headers.Count; i++)
                {
                    var hdr = data.Headers[i];
                    var val = row.StrList[i];

                    if (key != string.Empty && (hdr.FieldName == "id" || hdr.FieldName == "key"))
                        continue;

                    if (string.IsNullOrEmpty(val) && !(hdr.FieldType == "group" || hdr.FieldType == "string" || hdr.FieldType == "table"))
                        continue;

                    string s = string.Empty;
                    switch (hdr.FieldType)
                    {
                        case "string":
                            s = string.Format("\"{0}\"", val);
                            break;
                        case "string(nil)":
                            s = string.Format("\"{0}\"", val);
                            break;
                        case "int":
                            {
                                int.TryParse(val, out int n);
                                s = n.ToString();
                            }
                            break;
                        case "double":
                            {
                                double.TryParse(val, out double n);
                                s = n.ToString();
                            }
                            break;
                        case "bool":
                            {
                                bool.TryParse(val, out bool n);
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
                                s = BuildDataString(data.ChildData[hdr.FieldName], row.StrList[1], iDeep + 1);
                            }
                            break;
                    }
                    if (!string.IsNullOrEmpty(s))
                    {
                        if(bWithIndent == true)
                        {
                            for (int t = 1; t <= iDeep; t++)
                            {
                                csString.Append("    ");
                            }
                        }

                        if (hdr.FieldName == "id")
                            csString.AppendFormat("{0} = {1},", "index", s);
                        else if(hdr.FieldName == "key")
                            csString.AppendFormat("{0} = {1},", "id", s);
                        else
                            csString.AppendFormat("{0} = {1},", hdr.FieldName, s);
                        
                        if(bWithIndent == true)
                        {
                            csString.Append("\n");
                        }
                        else
                        {
                            csString.Append(" ");
                        }
                    }  
                }

                if(key != string.Empty && bWithIndent == true)
                {
                    for (int t = 1; t <= iDeep - 1; t++)
                    {
                        csString.Append("    ");
                    }
                }

                if (key == string.Empty)
                {
                    if (bWithIndent == true)
                    {
                        csString.AppendLine("    },\n");
                    }
                    else
                    {
                        csString.AppendLine("},");
                    }
                }
                else
                {
                    csString.Append("}");
                }
                    
                rbracket++;
            }
            if (lbracket != rbracket)
                Log.Err("key: {2} Bracket mismatch {0} to {1}", lbracket, rbracket, key);
            //else if (lbracket == 0)
            //    csString.Append("{}");

            return csString.ToString();
        }


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

            appendFormatLineEx(luaString, 0, "local indexItems = ");
            appendFormatLineEx(luaString, 0, "{{");
            for (int i = 0; i < ids.Count; i++)
            {
                appendFormatLineEx(luaString, 1, "[{0}] = items[{1}],", ids[i], i + 1);
            }
            appendFormatLineEx(luaString, 0, "}}");
            luaString.AppendLine();

            appendFormatLineEx(luaString, 0, "local idItems = ");
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
local data = { Items = items, IndexItems = indexItems, IdItems = idItems, }
");
            luaString.AppendFormat(
@"
function data:GetByIndex( index, prop )
    local item = self.IndexItems[index];
    if item == nil then
        sGlobal:Print( ""{0} GetByIndex nil item: ""..index );
        return index;
    end
    if prop == nil then
        return item;
    end
    if item[prop] == nil then
        sGlobal:Print( ""{0} GetByIndex nil prop: ""..prop );
        return item;
    end
    return item[prop];
end

function data:GetById( id, prop )
    local item = self.IdItems[id];
    if item == nil then
        sGlobal:Print( ""{0} GetById nil id: ""..id );
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

function data:GetIdByIndex( index )
    return data:GetByIndex( index, ""id"" );
end

function data:GetIndexById( id )
    return data:GetById( id, ""index"" );
end

function data:GetCount()
    return #self.IndexItems;
end

return data
", Path.GetFileName(fileName));
            return luaString.ToString();
        }
    }
}
