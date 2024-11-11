using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 考核系统.Entity;
using 考核系统.Mapper;
namespace 考核系统.Utils.CalcUtils
{
    internal class GlobalDeptInfo
    {
        private static List<Department> deptInfo;
        private static List<DeptAnnualInfo> deptAnnualInfo;
        private static GlobalDeptInfo instance;
        private GlobalDeptInfo()
        {
            var deptMapper = DepartmentMapper.GetInstance();
            var sql = "select * from department";
            deptInfo = deptMapper.QueryAll(sql);
            var deptAnnualInfoMapper = DeptAnnualInfoMapper.GetInstance();
            sql = "select * from dept_annual_info";
            deptAnnualInfo = deptAnnualInfoMapper.QueryAll(sql);
        }
        public static GlobalDeptInfo GetInstance()
        {
            if (instance == null)
            {
                instance = new GlobalDeptInfo();
            }
            return instance;
        }

        public int GlobalPopulation(int year)
        {
            //返回全校某年的总人数
            return Convert.ToInt32(deptAnnualInfo.Where(info => info.year == year).Sum(info => info.dept_population));
        }
    }
}
