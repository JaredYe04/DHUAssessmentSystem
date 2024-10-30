using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 考核系统.Utils
{
    public static class EventBus
    {
        //事件总线
        public delegate void YearChangedEventHandler(int year);
        public static event YearChangedEventHandler YearChanged;
        public static void OnYearChanged(int year)
        {
            YearChanged?.Invoke(year);
        }


    }
}
