using Microsoft.EntityFrameworkCore.Metadata.Internal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 考核系统.Entity
{
    class Index: DeepCopy<Index>//指标
    {
        public int id { get; set; }
        public int identifier_id { get; set; }
        public int secondary_identifier { get; set; }
        public string tertiary_identifier { get; set; }
        public string index_name { get; set; }
        public string index_type { get; set; }
        public double weight1 { get; set; }
        public double weight2 { get; set; }
        public double sensitivity { get; set; }

        public Index(int index_id, int identifier_id, int secondary_identifier,string tertiary_identifier, string index_name, string index_type, double weight1, double weight2, double sensitivity)
        {
            this.id = index_id;
            this.identifier_id = identifier_id;
            this.secondary_identifier = secondary_identifier;
            this.tertiary_identifier=tertiary_identifier;
            this.index_name = index_name;
            this.index_type = index_type;
            this.weight1 = weight1;
            this.weight2 = weight2;
            this.sensitivity = sensitivity;
        }
        public Index()
        {
        }
        public double BasicTheoreticalFullScore//5.单项基础类指标完成度理论满分
        {
            get
            {
                if (index_type == "加分类") return 0;
                return 100 * weight1 * weight2;
            }
        }
    }
}
