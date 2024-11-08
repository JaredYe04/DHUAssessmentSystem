using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 考核系统.Utils
{
    public enum DeptInfoColumns
    {
        id = 0,
        dept_code = 1,
        dept_name = 2,
        dept_population = 3,
        dept_punishment = 4,
        dept_group = 5
    }//部门信息表的列
    public enum IndexInfoColumns
    {
        id = 0,
        identifier_id = 1,
        secondary_identifier = 2,
        tertiary_identifier = 3,
        index_name = 4,
        index_type = 5,
        weight1 = 6,
        weight2 = 7,
        enable_sensitivity = 8,
        sensitivity = 9
    }//指标信息表的列
    public enum ManagerInfoColumns
    {
        id = 0,
        manager_code = 1,
        manager_name = 2
    }//职能部门信息表的列

    public enum IndexIdentifierInfoColumns
    {
        id = 0,
        identifier_name = 1
    }//指标分类信息表的列

    public enum CompletionColumns
    {
        id=0,
        dept_code = 1,
        dept_name = 2,
        target = 3,
        completed = 4,
        completion_rate = 5
    }
    public static class ColumnMap
    {

        public static Dictionary<DeptInfoColumns, string> deptInfoColumnsMap = new Dictionary<DeptInfoColumns, string>
        {
            {DeptInfoColumns.id,"单位编号"},
            {DeptInfoColumns.dept_code,"单位代码"},
            {DeptInfoColumns.dept_name,"单位名称"},
            {DeptInfoColumns.dept_population,"单位人数"},
            {DeptInfoColumns.dept_punishment,"惩罚系数"},
            {DeptInfoColumns.dept_group,"单位组别"}
        };

        public static Dictionary<IndexInfoColumns, string> indexInfoColumnsMap = new Dictionary<IndexInfoColumns, string>
        {
            {IndexInfoColumns.id,"指标编号"},
            {IndexInfoColumns.identifier_id,"一级类别"},
            {IndexInfoColumns.secondary_identifier,"二级类别"},
            {IndexInfoColumns.index_name,"指标名称"},
            {IndexInfoColumns.index_type,"指标类型"},
            {IndexInfoColumns.weight1,"权重1"},
            {IndexInfoColumns.weight2,"权重2"},
            {IndexInfoColumns.enable_sensitivity,"启用敏感度"},
            {IndexInfoColumns.sensitivity,"敏感度"}
        };


        public static Dictionary<ManagerInfoColumns, string> managerInfoColumnsMap = new Dictionary<ManagerInfoColumns, string>
        {
            {ManagerInfoColumns.id,"职能部门编号"},
            {ManagerInfoColumns.manager_code,"职能部门代码"},
            {ManagerInfoColumns.manager_name,"职能部门名称"}
        };
    }
}
