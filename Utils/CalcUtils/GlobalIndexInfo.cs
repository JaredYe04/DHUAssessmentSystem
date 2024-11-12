using Microsoft.EntityFrameworkCore.Metadata.Internal;
using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 考核系统.Entity;
using 考核系统.Mapper;
namespace 考核系统.Utils.CalcUtils
{
    internal class GlobalIndexInfo
    {
        private static List<Entity.Index> indexInfo;
        private static List<Completion> completionColumns;
        private static List<Department> departmentColumns;
        private static GlobalIndexInfo instance;
        private static List<GroupCompletion> groupCompletionColumns;
        private static int currentYear;
        private GlobalIndexInfo() {
            
            var indexMapper = IndexMapper.GetInstance();
            var sql= "select * from indexes";
            indexInfo = indexMapper.QueryAll(sql);
            var completionMapper = CompletionMapper.GetInstance();
            sql = "select * from completion";
            completionColumns = completionMapper.QueryAll(sql);
            var groupCompletionMapper = GroupCompletionMapper.GetInstance();
            sql = "select * from group_completion";
            groupCompletionColumns = groupCompletionMapper.QueryAll(sql);
            var deptMapper = DepartmentMapper.GetInstance();
            sql = "select * from department";
            departmentColumns = deptMapper.QueryAll(sql);
            currentYear = CommonData.CurrentYear;
        }
        public static GlobalIndexInfo GetInstance()
        {
            if (instance == null)
            {
                instance = new GlobalIndexInfo();
            }
            return instance;
        }

        public int GlobalCompletion(Entity.Index index)
        {
            //返回全校某指标的完成数之和
            int individualData = completionColumns.Where(completion => completion.index_id == index.id).Sum(completion => completion.completed);
            int groupData = groupCompletionColumns.Where
                (groupCompletion => groupCompletion.index_id == index.id
                && groupCompletion.year == currentYear).Sum(groupCompletion => groupCompletion.completed);
            var groups = GroupsMapper.GetInstance().GetGroupsByIndexId(index.id);
            int additionalData = 0;
            //要去除合并分组之前，各部门各自的完成数
            foreach (var completion in completionColumns)
            {
                var dept_id = completion.dept_id;
                var dept_code = departmentColumns.Where(dept => dept.id == dept_id).First().dept_code;
                //使用NaturalComparer类的Between方法判断dept_id是否在group的左右边界之间

                foreach (var group in groups)
                {
                    if (new NaturalComparer().Between(group.l_bound, group.r_bound, dept_code))
                    {
                        additionalData += completion.completed;
                        break;
                    }
                }
            }
            return individualData + groupData - additionalData;
        }
        public double BasicTheoreticalFullScoreSum//6.单项基础类指标完成度理论满分总分
        {
            get
            {
                return indexInfo.Where(index => index.index_type != "加分类").Sum(index => index.BasicTheoreticalFullScore);
            }
        }
    }
}
