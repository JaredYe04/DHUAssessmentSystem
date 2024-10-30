
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 考核系统.Utils;
using 考核系统.Entity;
namespace 考核系统.Mapper
{
    internal class ManagerMapper:BaseMapper<Manager>
    {
        private static ManagerMapper instance;
        private ManagerMapper() : base("manager")
        { }
        public static ManagerMapper GetInstance()
        {
            if (instance == null)
            {
                instance = new ManagerMapper();
            }
            return instance;
        }

    }
}
