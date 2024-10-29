using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 考核系统.Entity;

namespace 考核系统.Table
{
    class ResultTable//最后要输出的结果表
    {
        static Dictionary<String, Dictionary<String, ResultSet>> resultTable = new Dictionary<String, Dictionary<String, ResultSet>>();
        //resultTable的key是部门id，value是一个Dictionary，key是指标id,value是ResultSet


        public static void CalculateAll()
        {
            //遍历所有部门
            foreach (Department department in DepartmentTable.GetDepartmentTable().Values)
            {
                //遍历所有指标
                foreach (Index index in IndexTable.GetIndexTable().Values)
                {
                    //计算部门的指标得分
                    ResultSet.Calculate(department.Id, index.Id);
                }
            }
        }
    }
}
