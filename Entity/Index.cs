using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 考核系统.Entity
{
    class Index//指标
    {
        string id;//指标编号
        string name;//指标名称
        string type;//指标类型，加分类或基础类
        double firstWeight;//一级权重
        double secondWeight;//二级权重

        public string Id
        {
            get { return id; }
            set { id = value; }
        }

        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        public string Type
        {
            get { return type; }
            set { type = value; }
        }

        public double FirstWeight
        {
            get { return firstWeight; }
            set { firstWeight = value; }
        }

        public double SecondWeight
        {
            get { return secondWeight; }
            set { secondWeight = value; }
        }

        public Index(string id, string name, string type, double firstWeight, double secondWeight)
        {
            this.id = id;
            this.name = name;
            this.type = type;
            this.firstWeight = firstWeight;
            this.secondWeight = secondWeight;
        }
    }
}
