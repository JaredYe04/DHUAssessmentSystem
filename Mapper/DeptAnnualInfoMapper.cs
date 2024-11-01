using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 考核系统.Entity;
namespace 考核系统.Mapper
{
    internal class DeptAnnualInfoMapper: BaseMapper<DeptAnnualInfo>
    {
        private static DeptAnnualInfoMapper instance;
        private DeptAnnualInfoMapper() : base("dept_annual_info")
        { }
        public static DeptAnnualInfoMapper GetInstance()
        {
            if (instance == null)
            {
                instance = new DeptAnnualInfoMapper();
            }
            return instance;
        }
        public DeptAnnualInfo GetDeptAnnualInfo(int dept_id,int year)
        {
            var sql = $"select * from dept_annual_info where dept_id={dept_id} and year={year}";
            var result = instance.Query(sql);
            return result;
        }
    }

}
