using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 考核系统.Entity;

namespace 考核系统.Mapper
{
    internal class DepartmentMapper : BaseMapper<Department>
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

        public List<Department> GetDepartmentsByGroupName(string group_name, int year)
        {
            var deptAnnualInfoMapper = DeptAnnualInfoMapper.GetInstance();
            var deptAnnualInfos = deptAnnualInfoMapper.QueryAll
                ($"select * from dept_annual_info where year={year}");
            var deptIds = deptAnnualInfos.Where(x => x.dept_group == group_name).Select(x => x.dept_id).ToList();
            var sql = $"select * from department where id in ({string.Join(",", deptIds)})";
            return QueryAll(sql);
        }
    }
}
