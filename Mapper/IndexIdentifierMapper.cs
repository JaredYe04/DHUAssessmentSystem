using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 考核系统.Entity;

namespace 考核系统.Mapper
{
    internal class IndexIdentifierMapper:BaseMapper<IndexIdentifier>
    {
        public static IndexIdentifierMapper instance;
        private IndexIdentifierMapper() : base("index_identifier")
        { }
        public static IndexIdentifierMapper GetInstance()
        {
            if (instance == null)
            {
                instance = new IndexIdentifierMapper();
            }
            return instance;
        }
    }
}
