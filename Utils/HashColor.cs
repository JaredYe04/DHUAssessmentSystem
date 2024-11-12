using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 考核系统.Utils
{
    internal class HashColor
    {
        //为了区分分组，因此需要根据分组的名称生成一个颜色，这里使用了一个简单的哈希函数；生成的颜色应当是一个比较浅的颜色
        private static List<Color> colors = new List<Color>
        {
            Color.LightBlue,
            Color.LightCoral,
            Color.LightCyan,
            Color.LightGoldenrodYellow,
            Color.LightGreen,
            Color.LightGray,
            Color.LightPink,
            Color.LightSalmon,
            Color.LightSeaGreen,
            Color.LightSkyBlue,
            Color.LightSlateGray,
            Color.LightSteelBlue,
            Color.LightYellow
        };
        public static Color GetColor(string name)
        {
            // 使用哈希值生成种子值
            int hash = name.GetHashCode();

            // 生成随机数生成器
            Random random = new Random(new Random(hash).Next());

            int idx = random.Next(0, colors.Count);
            return colors[idx];

        }
    }
}
