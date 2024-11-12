using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 考核系统.Entity;
namespace 考核系统.Mapper
{
    internal class CompletionMapper : BaseMapper<Completion>
    {
        private static CompletionMapper instance;
        private CompletionMapper() : base("completion")
        { }
        public static CompletionMapper GetInstance()
        {
            if (instance == null)
            {
                instance = new CompletionMapper();
            }
            return instance;
        }

        public List<Completion> GetCompletionByIndexId(int id,int year)
        {
            var sql = $"select * from completion where index_id={id} and year={year}";
            return QueryAll(sql);
        }

        public List<Completion> GetIndexCompletionByYear(int currentYear)
        {
            var sql = $"select * from completion where year={currentYear}";
            return QueryAll(sql);
        }
        public void Update(Completion completion)
        {
            string[] bypassKeys = { "completion_rate" };
            base.Update(completion, bypassKeys);
        }
        public void Add(Completion completion, bool AutoId = true)
        {
            string[] bypassKeys = { "completion_rate" };
            base.Add(completion, AutoId, bypassKeys);
        }
    }

}
