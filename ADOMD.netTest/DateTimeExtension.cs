using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADOMD.netTest
{
    public static class DateTimeExtension
    {
       
        /// <summary>
        /// 获取上期 和 同期 时间段
        /// </summary>
        /// <param name="startTime"></param>
        /// <param name="endTime"></param>
        /// <returns></returns>
        public static Dictionary<string,DateTime> LastPeriodAndSamePeriod(DateTime startTime, DateTime endTime)
        {
            Dictionary<string, DateTime> dic = new Dictionary<string, DateTime>();
            if (endTime > startTime) {
                int days = (endTime - startTime).Days;
                var lp_et = startTime.AddDays(-1);
                var lp_st = lp_et.AddDays(-days);
                var sp_et = endTime.AddYears(-1);
                var sp_st = sp_et.AddDays(-days);
                dic.Add(DatePeriod.LP_ST.ToString(), lp_st);
                dic.Add(DatePeriod.LP_ET.ToString(), lp_et);
                dic.Add(DatePeriod.SP_ST.ToString(), sp_st);
                dic.Add(DatePeriod.SP_ET.ToString(), sp_et);
            }
            return dic;
        }

       
    }

    public enum DatePeriod
    {
        /// <summary>
        /// 上期 开始时间段
        /// </summary>
        LP_ST,
        /// <summary>
        /// 上期 结束时间段
        /// </summary>
        LP_ET,
        /// <summary>
        /// 同期 开始时间段
        /// </summary>
        SP_ST,
        /// <summary>
        /// 同期 结束时间段
        /// </summary>
        SP_ET
    }
}
