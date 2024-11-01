using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 考核系统.Entity
{
    class Department:DeepCopy<Department>//部门
    {
        [Key]
        public int id { get; set; }//部门编号
        public string dept_code { get; set; }//部门代码
        public string dept_name { get; set; }//部门名称

        public Department(int dept_id,string dept_code, string dept_name)
        {
            this.id = dept_id;
            this.dept_code = dept_code;
            this.dept_name = dept_name;
        }
        public Department()
        {
        }


    }

}