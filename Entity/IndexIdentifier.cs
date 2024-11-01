using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 考核系统.Entity
{
    internal class IndexIdentifier: DeepCopy<IndexIdentifier>//指标的一级编号
    {
        public int id { get; set; }
        public string identifier_name { get; set; }

        public IndexIdentifier(int identifier_id, string identifier_name)
        {
            this.id = identifier_id;
            this.identifier_name = identifier_name;
        }
        public IndexIdentifier()
        {
        }
    }
}
