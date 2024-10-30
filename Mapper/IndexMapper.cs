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
        private IndexMapper() : base("index")
        { }
        public static IndexMapper GetInstance()
        {
            if (instance == null)
            {
                instance = new IndexMapper();
            }
            return instance;
        }
    }

}
