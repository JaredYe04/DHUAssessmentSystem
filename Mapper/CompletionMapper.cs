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

        public List<Completion> GetCompletionByIndexId(int id)
        {
            var sql = $"select * from completion where index_id={id}";
            return QueryAll(sql);
        }

        public List<Completion> GetIndexCompletionByYear(int currentYear)
        {
            var sql = $"select * from completion where year={currentYear}";
            return QueryAll(sql);
        }
    }

}
