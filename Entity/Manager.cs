using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 考核系统.Entity
{
    public class Manager
    {
        public int id { get; set; }
        public string manager_code { get; set; }
        public string manager_name { get; set; }

        public Manager(int managerId,string manager_code, string managerName)
        {
            id = managerId;
            this.manager_code = manager_code;
            manager_name = managerName;
        }
    }
}
