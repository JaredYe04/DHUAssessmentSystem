using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 考核系统.Entity
{
    class DeptAnnualInfo: DeepCopy<DeptAnnualInfo>//部门年度信息
    {
        public int id { get; set; }
        public int dept_id { get; set; }
        public int year { get; set; }
        public int dept_population { get; set; }
        public double dept_punishment { get; set; }
        public string dept_group { get; set; }

        public DeptAnnualInfo(int id, int dept_id, int year, int dept_population, double dept_punishment, string dept_group)
        {
            this.id = id;
            this.dept_id = dept_id;
            this.year = year;
            this.dept_population = dept_population;
            this.dept_punishment = dept_punishment;
            this.dept_group = dept_group;
        }
        public DeptAnnualInfo(int dept_id, int year, int dept_population, double dept_punishment, string dept_group)
        {
            this.dept_id = dept_id;
            this.year = year;
            this.dept_population = dept_population;
            this.dept_punishment = dept_punishment;
            this.dept_group = dept_group;
        }//id自增
        public DeptAnnualInfo()
        {
        }
    }
}
