using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

public class NaturalComparer : IComparer<string>
{
    public int Compare(string x, string y)
    {
        if (x == y) return 0;
        if (x == null) return -1;
        if (y == null) return 1;

        var regex = new Regex(@"(\d+|\D+)");
        var xParts = regex.Matches(x).Cast<Match>().Select(m => m.Value).ToArray();
        var yParts = regex.Matches(y).Cast<Match>().Select(m => m.Value).ToArray();

        for (int i = 0; i < Math.Min(xParts.Length, yParts.Length); i++)
        {
            if (xParts[i] != yParts[i])
            {
                if (int.TryParse(xParts[i], out int xNum) && int.TryParse(yParts[i], out int yNum))
                {
                    return xNum.CompareTo(yNum);
                }
                return xParts[i].CompareTo(yParts[i]);
            }
        }
        return xParts.Length.CompareTo(yParts.Length);
    }
    //Between函数，输入l_bound, r_bound, x，根据Compare来比较字符串，返回l_bound <= x <= r_bound
    public bool Between(string l_bound, string r_bound, string x)
    {
        return Compare(l_bound, x) <= 0 && Compare(x, r_bound) <= 0;
    }
}
