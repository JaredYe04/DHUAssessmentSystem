using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 考核系统.Entity;

namespace 考核系统.Mapper
{
    internal class IndexMapper : BaseMapper<Index>
    {
        private static IndexMapper instance;
        private IndexMapper() : base("indexes")
        { }
        public static IndexMapper GetInstance()
        {
            if (instance == null)
            {
                instance = new IndexMapper();
            }
            return instance;
        }
        //重写Update方法
        public void Update(Index index)
        {
            string[] bypassKeys = { "BasicTheoreticalFullScore" };
            base.Update(index, bypassKeys);
        }
        public void Add(Index index, bool AutoId = true)
        {
            string[] bypassKeys = { "BasicTheoreticalFullScore" };
            base.Add(index, AutoId, bypassKeys);
        }
    }

}
