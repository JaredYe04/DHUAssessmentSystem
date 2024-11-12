using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 考核系统.Entity;
namespace 考核系统.Mapper
{
    internal class GroupCompletionMapper : BaseMapper<GroupCompletion>
    {
        private static GroupCompletionMapper instance;
        private GroupCompletionMapper() : base("group_completion")
        { }
        public static GroupCompletionMapper GetInstance()
        {
            if (instance == null)
            {
                instance = new GroupCompletionMapper();
            }
            return instance;
        }

        public List<GroupCompletion> GetCompletionByIndexId(int id,int year)
        {
            var sql = $"select * from group_completion where index_id={id} and year={year}";
            return QueryAll(sql);
        }

        public List<GroupCompletion> GetIndexCompletionByYear(int currentYear)
        {
            var sql = $"select * from group_completion where year={currentYear}";
            return QueryAll(sql);
        }
        public void Update(GroupCompletion completion)
        {
            string[] bypassKeys = { "completion_rate" };
            base.Update(completion, bypassKeys);
        }
        public void Add(GroupCompletion completion, bool AutoId = true)
        {
            string[] bypassKeys = { "completion_rate" };
            base.Add(completion, AutoId, bypassKeys);
        }
    }

}
