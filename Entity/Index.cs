using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 考核系统.Entity
{
    public class Index
    {
        public int id { get; set; }
        public string index_code { get; set; }
        public string index_name { get; set; }
        public string index_type { get; set; }
        public double weight1 { get; set; }
        public double weight2 { get; set; }
        public int enable_sensitivity { get; set; }
        public double sensitivity { get; set; }

        public Index(int index_id,string index_code, string index_name, string index_type, double weight1, double weight2, int enable_sensitivity, double sensitivity)
        {
            this.id = index_id;
            this.index_code = index_code;
            this.index_name = index_name;
            this.index_type = index_type;
            this.weight1 = weight1;
            this.weight2 = weight2;
            this.enable_sensitivity = enable_sensitivity;
            this.sensitivity = sensitivity;
        }
    }
}
