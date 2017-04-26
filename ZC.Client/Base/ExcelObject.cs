using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ZC.Client.Base
{
    public class ExcelObject
    {
        public string MachineId { get; set; }
        public int TotalCheck { get; set; }
        public int TotalGood { get; set; }
        public double TotalGoodPercent { get; set; }
        public int TotalBad { get; set; }
        public double TotalBadPercent { get; set; }

    }
}
