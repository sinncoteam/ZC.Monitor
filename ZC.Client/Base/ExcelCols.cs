using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ZC.Client.Base
{
    public enum ExcelCols
    {
        /// <summary>
        /// B列，检测日期
        /// </summary>
        检测日期 = 1,
        /// <summary>
        /// F列，检测机号
        /// </summary>
        检测机号 = 5,
        /// <summary>
        /// J列，检验总数
        /// </summary>
        检验总数 = 9,
        /// <summary>
        /// K列，合格总数
        /// </summary>
        总合格数 = 10,
        /// <summary>
        /// L列，合格率
        /// </summary>
        总合格率 = 11,
        /// <summary>
        /// M列，总不良数
        /// </summary>
        总不良数 = 12,
        /// <summary>
        /// N列，总不良率
        /// </summary>
        总不良率 = 13
    }
}
