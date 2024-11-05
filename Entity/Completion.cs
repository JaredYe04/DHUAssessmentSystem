using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 考核系统.Entity
{
    internal class Completion: DeepCopy<Completion>
    {
        public int id { get; set; }
        public int dept_id { get; set; }
        public int year { get; set; }
        public int index_id { get; set; }
        public int target { get; set; }
        public int completed { get; set; }

        public Completion(int completion_id, int dept_id, int year, int index_id, int target, int completed)
        {
            this.id = completion_id;
            this.dept_id = dept_id;
            this.year = year;
            this.index_id = index_id;
            this.target = target;
            this.completed = completed;
        }
        public Completion() { }
    }
}
