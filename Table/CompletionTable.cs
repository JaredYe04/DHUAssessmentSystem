using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 考核系统.Entity;

namespace 考核系统.Table
{
    class CompletionTable
    {
        static Dictionary<string, Dictionary<string, CompletionSet>> completionTable = new Dictionary<string, Dictionary<string, CompletionSet>>();



        public static void AddCompletionSet(string departmentId, string indexId, CompletionSet completionSet)
        {
            if (!completionTable.ContainsKey(departmentId))
            {
                completionTable.Add(departmentId, new Dictionary<string, CompletionSet>());
            }
            completionTable[departmentId].Add(indexId, completionSet);
        }
        public static CompletionSet GetCompletionSet(string departmentId, string indexId)
        {
            return completionTable[departmentId][indexId];
        }
        public static void RemoveCompletionSet(string departmentId, string indexId)
        {
            completionTable[departmentId].Remove(indexId);
        }
        public static void UpdateCompletionSet(string departmentId, string indexId, CompletionSet completionSet)
        {
            completionTable[departmentId][indexId] = completionSet;
        }

        public static List<Tuple<string, CompletionSet>> getCompletionSetByDepartment(string departmentId)//返回某个部门的所有指标的完成情况
        {
            List<Tuple<string, CompletionSet>> result = new List<Tuple<string, CompletionSet>>();
            foreach (var item in completionTable[departmentId])
            {
                result.Add(new Tuple<string, CompletionSet>(item.Key, item.Value));
            }
            return result;
        }

        public static List<Tuple<string,CompletionSet>> getCompletionSetByIndex(string indexId)//返回所有部门的某个指标的完成情况
        {
            List<Tuple<string, CompletionSet>> result = new List<Tuple<string, CompletionSet>>();
            foreach (var item in completionTable)
            {
                if (item.Value.ContainsKey(indexId))
                {
                    result.Add(new Tuple<string, CompletionSet>(item.Key, item.Value[indexId]));
                }
            }
            return result;
        }

        public static string Serialize()
        {
            //newtonsoft.json.jsonconvert.serializeobject
            return Newtonsoft.Json.JsonConvert.SerializeObject(completionTable);
        }
        public static void Parse(string Json)
        {
            //newtonsoft.json.jsonconvert.deserializeobject
            completionTable = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, CompletionSet>>>(Json);
        }

    }
}
