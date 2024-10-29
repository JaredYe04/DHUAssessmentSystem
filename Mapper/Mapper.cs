using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 考核系统.Entity;

namespace 考核系统.Mapper
{
    class Mapper
    {
        public static Index GetIndexById(string id)
        {
            return Table.IndexTable.GetIndexById(id);
        }
        public static Department GetDepartmentById(string id)
        {
            return Table.DepartmentTable.GetDepartmentById(id);
        }
        public static Index GetIndexByName(string name)
        {
            Dictionary<String, Index> indexTable = Table.IndexTable.GetIndexTable();
            foreach (Index index in indexTable.Values)
            {
                if (index.Name == name)
                {
                    return index;
                }
            }
            throw new Exception("找不到该指标:"+name);
        }
        public static Department GetDepartmentByName(string name)
        {
            Dictionary<String, Department> departmentTable = Table.DepartmentTable.GetDepartmentTable();
            foreach (Department department in departmentTable.Values)
            {
                if (department.Name == name)
                {
                    return department;
                }
            }
            throw new Exception("找不到该部门:"+name);
        }
    }
}
