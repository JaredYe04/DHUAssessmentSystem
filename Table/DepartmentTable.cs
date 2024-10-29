using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 考核系统.Entity;
namespace 考核系统.Table
{
    static class DepartmentTable//部门信息，需要自动解析或者由Excel导入
    {
        static Dictionary<String, Department> departmentTable = new Dictionary<String, Department>();
        public static void AddDepartment(Department department)
        {
            departmentTable.Add(department.Id, department);
        }
        public static Department GetDepartmentById(string id)
        {
            return departmentTable[id];
        }
        public static void RemoveDepartmentById(string id)
        {
            departmentTable.Remove(id);
        }
        public static void UpdateDepartmentById(string id, Department department)
        {
            departmentTable[id] = department;
        }
        public static Dictionary<String, Department> GetDepartmentTable()
        {
            return departmentTable;
        }

        public static string Serialize()
        {
            //使用Newtonsoft.Json.JsonConvert.SerializeObject方法序列化对象
            return Newtonsoft.Json.JsonConvert.SerializeObject(departmentTable);
        }
        public static void Parse(string Json)
        {
            //使用Newtonsoft.Json.JsonConvert.DeserializeObject方法反序列化对象
            departmentTable = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<String, Department>>(Json);
        }
    }
}
