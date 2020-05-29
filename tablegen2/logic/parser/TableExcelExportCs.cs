using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace tablegen2.logic
{
    public static class TableExcelExportCs
    {
        public static void ExportExcelFile(TableExcelData data, string filePath)
        {
            var csString = new StringBuilder();

            csString.AppendPrologue();
            csString.AppendTitle(filePath);
            csString.AppendProperty(data, string.Empty, 1);
            csString.AppendEnd();

            File.WriteAllBytes(filePath, Encoding.UTF8.GetBytes(csString.ToString()));
        }

        //-- 构造开头 ---------------------------------------------------------------------------------------------------------------------------
        private static void AppendPrologue(this StringBuilder sb)
        {
            sb.Append(
@"
using System.Collections.Generic;

namespace Feamber.Data
{
");
        }

        //-- 构造类名 ---------------------------------------------------------------------------------------------------------------------------
        private static void AppendTitle(this StringBuilder sb, string path)
        {
            sb.AppendLine($"    public sealed class {Path.GetFileNameWithoutExtension(path)} : IData");
            sb.AppendLine("    {");
        }

        //-- 构造成员 ---------------------------------------------------------------------------------------------------------------------------
        private static void AppendProperty(this StringBuilder sb, TableExcelData data, string key, int deep)
        {
            for (int i = 0; i < data.Headers.Count; i++)
            {
                var header = data.Headers[i];
                
                if (key != string.Empty && (header.FieldName.ToLower() == "id" || header.FieldName.ToLower() == "key"))
                    continue;

                string type = header.FieldType;
                string name = header.FieldName;
                string desc = header.FieldDesc;

                if (desc != string.Empty)
                {
                    if (name.ToLower() == "id") { desc = "唯一数字索引 (1~N)"; }
                    else if (name.ToLower() == "key") { desc = "唯一字符串索引"; }
                    sb.AppendComment(desc, deep);
                }

                if (type.Contains("group"))
                {
                    type = GetGroupType(type, data.Rows[0].StrList[i]);
                    sb.AppendListProperty(type, name, deep);
                }
                else if (type == "table")
                {
                    sb.AppendNestedClassTitle(name, deep);
                    sb.AppendProperty(data.ChildData[name], data.Rows[1].StrList[1], deep + 1);
                    sb.AppendNestedClassEnd(deep);
                }
                else
                {
                    if (name.ToLower() == "id") { name = "Index"; }
                    if (name.ToLower() == "key") { name = "Id"; }

                    if (type.ToLower() == "string(nil)") { type = "string"; }
                    if (type.ToLower() == "double") { type = "float"; }
                    if (type.ToLower() == "color") { type = "string"; }
                    sb.AppendNormalProperty(type, name, deep);
                }
                if (i != data.Headers.Count-1)
                {
                    sb.AppendLine();
                }
            }
        }


        //-- 构造结尾 ---------------------------------------------------------------------------------------------------------------------------
        private static void AppendEnd(this StringBuilder sb)
        {
            sb.AppendLine("    }");
            sb.Append("}");
        }

        //-- Append Property -------------------------------------------------------------------------------------------------------------------------
        private static void AppendNormalProperty(this StringBuilder sb, string type, string name, int deep)
        {
            sb.AppendIndent(deep);
            sb.AppendLine($"public {type} {name} {{ get; set; }}");
        }

        private static void AppendListProperty(this StringBuilder sb, string type, string name, int deep)
        {
            sb.AppendIndent(deep);
            sb.AppendLine($"public List<{type}> {name} {{ get; set; }}");
        }

        private static void AppendNestedClassTitle(this StringBuilder sb, string name, int deep)
        {
            sb.AppendIndent(deep);
            sb.AppendLine($"public Type_{name} {name} {{ get; set; }}");
            sb.AppendLine();
            sb.AppendIndent(deep);
            sb.AppendLine($"public class Type_{name}");
            sb.AppendIndent(deep);
            sb.AppendLine("{");
        }

        private static void AppendNestedClassEnd(this StringBuilder sb, int deep)
        {
            sb.AppendIndent(deep);
            sb.AppendLine("}");
        }

        private static void AppendComment(this StringBuilder sb, string comment, int deep)
        {
            var commentList = comment.Split('\n');
            foreach (var c in commentList)
            {
                sb.AppendIndent(deep);
                sb.AppendLine($"// {c}");
            }

            //complex
            //var commentList = comment.Split('\n');
            //sb.AppendIndent(deep);
            //sb.AppendLine($"/// <summary>");
            //foreach (var c in commentList)
            //{
            //    sb.AppendIndent(deep);
            //   sb.AppendLine($"/// {c}");
            //}
            //sb.AppendIndent(deep);
            //sb.AppendLine($"/// <summary>");
        }

        private static void AppendIndent(this StringBuilder sb, int deep)
        {
            for (int t = 0; t <= deep; t++)
            {
                sb.Append("    ");
            }
        }

        private static string GetGroupType(string type, string value)
        {
            var item = value.Split(',')[0];

            string t;
            if (type.Contains("int") || int.TryParse(item, out _))
            {
                t = "int";
            }
            else if (type.Contains("double") || double.TryParse(item, out _))
            {
                t = "float";
            }
            else if (type.Contains("bool") || bool.TryParse(item, out _))
            {
                t = "bool";
            }
            else
            {
                t = "string";
            }
            return t;
        }
    }
}