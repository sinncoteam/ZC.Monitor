using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ZC.Utils
{
    public class StringHelper
    {
        public static double PercentToInt(string percent)
        {
            if(string.IsNullOrEmpty(percent))
            {
                return 0;
            }
            percent = percent.TrimEnd('%');
            double per = Convert.ToDouble(percent);
            return per;
        }
    }
}
