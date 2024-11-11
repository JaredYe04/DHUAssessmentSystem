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
        public string group_name { get; set; }

        public Groups(int group_id, string group_name)
        {
            this.id = group_id;
            this.group_name = group_name;
        }
        public Groups() { }
    }
}
