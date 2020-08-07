using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace tablegen2.logic
{
    public static class TableExcelExportCsEx
    {
        public static void ExportExcelFile(TableExcelData data, string filePath, string fileName)
        {
            var csString = new StringBuilder();
            csString.AppendTemplate(fileName);

            File.WriteAllBytes(filePath, Encoding.UTF8.GetBytes(csString.ToString()));
        }

        private static void AppendTemplate(this StringBuilder sb, string className)
        {
            sb.AppendLine(@"using System.Collections.Generic;");
            sb.AppendLine(@"using UnityEngine;");
            sb.AppendLine(@"");
            sb.AppendLine(@"namespace Feamber.Data");
            sb.AppendLine(@"{");
            sb.AppendLine($"    public sealed partial class { className } : IBaseData");
            sb.AppendLine(@"    {");
            sb.AppendLine($"        public void OnValidate()");
            sb.AppendLine(@"        {");
            sb.AppendLine(@"            Debug.Assert(Index > 0, $""{ nameof(Index)}{ Index} (excel-id) must greater than 0"");");
            sb.AppendLine(@"        }");
            sb.AppendLine(@"    }");
            sb.AppendLine(@"");
            sb.AppendLine(@"    public static partial class DataListExtensioMethods");
            sb.AppendLine(@"    {");
            sb.AppendLine($"        public static void Demo(this DataList<{ className }> dataList)");
            sb.AppendLine(@"        {");
            sb.AppendLine(@"        }");
            sb.AppendLine(@"    }");
            sb.AppendLine(@"}");
            sb.AppendLine(@"");
        }
    }
}