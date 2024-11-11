using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 考核系统.Entity;

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

    }
}
