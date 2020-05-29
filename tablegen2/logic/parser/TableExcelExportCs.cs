using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

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
            csString.Append( BuildPrologue() );
            csString.Append( BuildTitle(filePath) );
            // appendFormatLineEx(csString, 0, "local items =");
            // appendFormatLineEx(csString, 0, "{{");

            csString.Append(BuildDataString(data, string.Empty, 1));

            //appendFormatLineEx(csString, 0, "}}");
            //csString.AppendLine();
            //csString.Append(BuildItemString(data));
            //csString.Append(BuildFuncString(Path.GetFileName(filePath)));
            csString.Append(BuildEnd());
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

//-- 构造开头 ---------------------------------------------------------------------------------------------------------------------------
        private static string BuildPrologue()
        {
            var prologue = new StringBuilder();
            prologue.Append(
@"
using System.Collections.Generic;

namespace Feamber.Data
{
");
            return prologue.ToString();
        }




 //-- 构造类名 ---------------------------------------------------------------------------------------------------------------------------
        private static string BuildTitle(string path)
        {
            var title = new StringBuilder();
            title.AppendLine($"    public sealed class {Path.GetFileNameWithoutExtension(path)} : IData");
            title.AppendLine("    {");
            return title.ToString();
        }

//-- 构造成员 ---------------------------------------------------------------------------------------------------------------------------
        private static string BuildDataString(TableExcelData data, string key, int deep)
        {

            var csString = new StringBuilder();

                for (int i = 0; i < data.Headers.Count; i++)
                {
                    var header = data.Headers[i];
                    //var value = data.Rows[1].StrList[i];

                    //if (key != string.Empty && (header.FieldName == "id" || header.FieldName == "key"))
                    //    continue;

                    //if (string.IsNullOrEmpty(value) && !(header.FieldType == "group" || header.FieldType == "string" || header.FieldType == "table"))
                    //    continue;

                    string name = string.Empty;
                    switch (header.FieldType)
                    {
                    case "string":
                    case "string(nil)":
                            name = header.FieldName;
                            if (header.FieldName.ToLower() == "key")
                            {
                                name = "Id";
                            }
                            csString.AppendNormalProperty("string", name, deep);
                            break;
                    case "int":

                            name = header.FieldName;
                            if (header.FieldName.ToLower() == "id")
                            {
                                name = "Index";
                            }
                            csString.AppendNormalProperty("int", name, deep);
                            break;
                        case "double":
                            csString.AppendNormalProperty("double", header.FieldName, deep);
                            break;
                        case "bool":
                        csString.AppendNormalProperty("bool", header.FieldName, deep);
                            {
                               // bool.TryParse(value, out bool n);
                              //  s = n == true ? "true" : "false";
                            }
                            break;
                        case "group":
                            {
                              //  s = string.Format("{{{0}}}", value);
                            }
                            break;
                        case "color":
                            {
                             //   s = string.Format("{0:X}", value);
                            }
                            break;
                        case "table":
                            {
                            //    s = BuildDataString(data.ChildData[header.FieldName], data.Rows[1].StrList[1], deep + 1);
                            }
                            break;
                    }
                }

            return csString.ToString();
        }


//-- 构造结尾 ---------------------------------------------------------------------------------------------------------------------------
        private static string BuildEnd()
        {
            var end = new StringBuilder();
            _ = end.Append(
@"
    }
}
");
            return end.ToString();
        }



        public static void AppendIndent(this StringBuilder sb, int deep)
        {
            for (int t = 0; t <= deep; t++)
            {
                sb.Append("    ");
            }
        }

        //-- Append Property -------------------------------------------------------------------------------------------------------------------------
        private static void AppendNormalProperty(this StringBuilder sb,string type, string name, int deep)
        {
            sb.AppendIndent(deep);
            sb.AppendLine($"public {type} {name} {{ get; set; }}");
            sb.AppendLine("");
        }

        private static void AppendListProperty(this StringBuilder sb, string type, string name, int deep)
        {

        }

        private static void AppendClassProperty(this StringBuilder sb, string name, int deep)
        {

        }















//-------------------------------------------------------------------------------------------------------------------
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
