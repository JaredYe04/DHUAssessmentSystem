using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 考核系统.Entity
{
    class Department//部门
    {
        string id;//部门编号
        string name;//部门名称
        int population;//部门人数
        double punishRatio;//扣分比例


        //get set constructor
        public string Id
        {
            get { return id; }
            set { id = value; }
        }
        public string Name
        {
            get { return name; }
            set
            {
                name = value;
            }

        }
        public int Population
        {
            get { return population; }
            set { population = value; }
        }
        public double PunishRatio
        {
            get { return punishRatio; }
            set { punishRatio = value; }
        }

        public Department(string id, string name, int population, double punishRatio)
        {
            this.id = id;
            this.name = name;
            this.population = population;
            this.punishRatio = punishRatio;
        }
    }

}