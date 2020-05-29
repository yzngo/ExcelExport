
using System.Collections.Generic;

namespace Feamber.Data
{
    public sealed class Student : IData
    {
        // 唯一数字索引 (1~N)
        public int Index { get; set; }

        // 唯一字符串索引
        public string Id { get; set; }

        // 年龄
        public int Age { get; set; }

        // 姓名
        public string Name { get; set; }

        // 体重
        public float Weight { get; set; }

        // 肤色
        public string Skin { get; set; }

        // 会说的语言
        public List<string> Language { get; set; }

        // 数字数组
        public List<int> Number { get; set; }

        // 小数数组
        public List<double> Score { get; set; }

        // 父母信息
        // 姓名等
        public Type_Parent Parent { get; set; }

        public class Type_Parent
        {
            // 姓名
            public string Name { get; set; }
        }
    }
}