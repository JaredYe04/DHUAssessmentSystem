using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 考核系统.Entity;

namespace 考核系统.Mapper
{
    internal class DepartmentMapper: BaseMapper<Department>
    {
        private static DepartmentMapper instance;
        private DepartmentMapper() : base("department")
        { }
        public static DepartmentMapper GetInstance()
        {
            if (instance == null)
            {
                instance = new DepartmentMapper();
            }
            return instance;
        }
    }
}
