using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 考核系统.Entity
{
    internal class GroupCompletion: DeepCopy<GroupCompletion>
    {
        public int id { get; set; }
        public int group_id { get; set; }
        public int year { get; set; }
        public int index_id { get; set; }
        public int target { get; set; }
        public int completed { get; set; }

        public GroupCompletion(int completion_id, int group_id, int year, int index_id, int target, int completed)
        {
            this.id = completion_id;
            this.group_id = group_id;
            this.year = year;
            this.index_id = index_id;
            this.target = target;
            this.completed = completed;
        }
        public GroupCompletion() { }
        public double completion_rate
        {
            get
            {
                if (target == 0) return 0;//防止除0
                return (double)completed / target;
            }
        }
    }
}
