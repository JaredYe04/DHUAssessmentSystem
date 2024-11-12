using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 考核系统.Entity;
using 考核系统.Utils;
namespace 考核系统.Mapper
{
    internal class GroupsMapper:BaseMapper<Groups>
    {
        private static GroupsMapper instance;
        private GroupsMapper() : base("groups")
        { }
        public static GroupsMapper GetInstance()
        {
            if (instance == null)
            {
                instance = new GroupsMapper();
            }
            return instance;
        }
        public Groups GetGroupByDeptCode(string deptCode,int index_id)
        {
            var list = GetAllObjects().Where(x=>x.index_id==index_id);
            //var list = GetAllObjects();
            var comparer = new NaturalComparer();
            foreach (var item in list)
            {
                if (comparer.Between(item.l_bound, item.r_bound, deptCode)){
                    return item;
                }
            }
            return null;
        }
        public List<Groups> GetGroupsByIndexId(int index_id)
        {
            var sql = $"select * from groups where index_id={index_id}";
            return QueryAll(sql);
        }
    }
}
