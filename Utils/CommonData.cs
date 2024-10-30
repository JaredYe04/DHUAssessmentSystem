using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 考核系统.Utils
{
    static class CommonData
    {
        private static int currentYear=DateTime.Now.Year;
        public static int CurrentYear
        {
            get { return currentYear; }
            set {
                currentYear = value;
                //向事件总线发送年份变更事件
                EventBus.OnYearChanged(currentYear);
            }
        }
    }
}
