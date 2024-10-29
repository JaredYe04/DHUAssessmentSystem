using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 考核系统.Entity;
namespace 考核系统.Table
{
    static class IndexTable//指标信息，见Sheet1
    {
        static Dictionary<String, Index> indexTable = new Dictionary<String, Index>();
        public static void AddIndex(Index index)
        {
            indexTable.Add(index.Id, index);
        }
        public static Index GetIndexById(string id)
        {
            return indexTable[id];
        }
        public static void RemoveIndexById(string id)
        {
            indexTable.Remove(id);
        }
        public static void UpdateIndexById(string id, Index index)
        {
            indexTable[id] = index;
        }
        public static Dictionary<String, Index> GetIndexTable()
        {
            return indexTable;
        }

        public static string Serialize()
        {
            //使用Newtonsoft.Json.JsonConvert.SerializeObject方法序列化对象
            return Newtonsoft.Json.JsonConvert.SerializeObject(indexTable);
        }
        public static void Parse(string Json)
        {
            //使用Newtonsoft.Json.JsonConvert.DeserializeObject方法反序列化对象
            indexTable = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<String, Index>>(Json);
        }
    }
}