using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 考核系统.Entity;
namespace 考核系统.Mapper
{
    internal class IndexDutyMapper : BaseMapper<IndexDuty>
    {
        private static IndexDutyMapper instance;
        private IndexDutyMapper() : base("index_duty")
        { }
        public static IndexDutyMapper GetInstance()
        {
            if (instance == null)
            {
                instance = new IndexDutyMapper();
            }
            return instance;
        }
    }
}
