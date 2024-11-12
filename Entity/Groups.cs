using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 考核系统.Entity
{
    internal class Groups:DeepCopy<Groups>
    {
        public int id { get; set; }
        public int index_id { get; set; }
        public string group_name { get; set; }
        public string l_bound { get; set; }
        public string r_bound { get; set; }

        public Groups(int group_id,int index_id, string group_name, string l_bound, string r_bound)
        {
            this.id = group_id;
            this.index_id = index_id;
            this.group_name = group_name;
            this.l_bound = l_bound;
            this.r_bound = r_bound;
        }
        public Groups() { }
    }
}
