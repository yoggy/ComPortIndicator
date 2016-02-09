using System;
using System.Text.RegularExpressions;

namespace ComPortIndicator
{
    class ComPortComparer : System.Collections.IComparer
    {
        Regex reg = new Regex("COM(?<num>\\d+)");

        public int Compare(object a, object b)
        {
            int a_num = int.Parse(reg.Match((string)a).Groups["num"].Value);
            int b_num = int.Parse(reg.Match((string)b).Groups["num"].Value);

            if (a_num == b_num)
            {
                return 0;
            }
            if (a_num < b_num)
            {
                return -1;
            }

            return 1;
        }
    }
}
