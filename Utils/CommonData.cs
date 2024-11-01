using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using 考核系统.Entity;

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
        public static Dictionary<int,Tuple<Department,DeptAnnualInfo>> DeptInfo { get; set; }
        public static Dictionary<int, Index> IndexInfo { get; set; }
        public static Dictionary<int, Manager> ManagerInfo { get; set; }

        public static Dictionary<int,IndexDuty>DutyInfo { get; set; }
        public static Dictionary<int, Index> UnallocatedIndexes
        {
            get
            {
                //在IndexInfo中，但不在DutyInfo中的指标
                Dictionary<int, Index> unallocatedIndex = new Dictionary<int, Index>();
                foreach (var index in IndexInfo.Values)
                {
                    if (!DutyInfo.Values.Any(duty => duty.index_id == index.id))
                    {
                        unallocatedIndex.Add(index.id, index);
                    }
                }
                return unallocatedIndex;
            }
        }
        public static Manager selectedManager { get; set; }

    }
}
