using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 考核系统.Entity
{
    public class IndexDuty
    {
        public int id { get; set; }
        public int manager_id { get; set; }
        public int index_id { get; set; }
        public int enable_assessment { get; set; }

        public IndexDuty(int index_duty_id, int managerId, int indexId, int enable_assessment)
        {
            this.id = index_duty_id;
            manager_id = managerId;
            index_id = indexId;
            this.enable_assessment = enable_assessment;
        }
    }
}
