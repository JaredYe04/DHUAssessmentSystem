using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 考核系统.Entity;
using 考核系统.Utils;

namespace 考核系统.Mapper
{
    internal class BaseMapper<T> where T : class
    {
        private static string tableName;
        private static string keyName;
        protected BaseMapper(string tableName, string keyName="id")
        {
            BaseMapper<T>.tableName = tableName;
            BaseMapper<T>.keyName = keyName;
        }
        public T Query(string sql)
        {
            var result = DB.GetInstance().ExecuteReader(sql)[0];
            if(result == null)
            {
                return null;
            }   
            string json = JsonConvert.SerializeObject(result);
            return JsonConvert.DeserializeObject<T>(json);
        }
        public List<T> QueryAll(string sql)
        {
            var results = DB.GetInstance().ExecuteReader(sql);
            List<T> list = new List<T>();
            foreach (var result in results)
            {
                string json = JsonConvert.SerializeObject(result);
                list.Add(JsonConvert.DeserializeObject<T>(json));
            }
            return list;
        }
        public List<T> GetAllObjects()
        {

            string sql = $"select * from '{tableName}'";
            return QueryAll(sql);
        }
        public void Add(T obj,bool AutoId=true)
        {
            try
            {
                string json = JsonConvert.SerializeObject(obj);
                Dictionary<string, object> dict = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                if(AutoId)
                {
                    dict.Remove(keyName);
                }
                string columns = string.Join(",", dict.Keys);
                string values = string.Join(",", dict.Values.Select(x => $"'{x}'"));
                string sql = $"insert into {tableName}({columns}) values({values})";
                DB.GetInstance().ExecuteNonQuery(sql);
            }
            catch (Exception e)
            {
                Logger.Log(e.Message, LogType.ERROR);
            }

        }
        public void Remove(string key)
        {
            try
            {
                string sql = $"delete from {tableName} where {keyName} = '{key}'";
                DB.GetInstance().ExecuteNonQuery(sql);
            }
            catch (Exception e)
            {
                Logger.Log(e.Message, LogType.ERROR);
            }
        }
        public void Remove(T obj)
        {
            try
            {
                string json = JsonConvert.SerializeObject(obj);
                Dictionary<string, object> dict = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                string sql = $"delete from {tableName} where {keyName} = '{dict[keyName]}'";
                DB.GetInstance().ExecuteNonQuery(sql);
            }
            catch (Exception e)
            {
                Logger.Log(e.Message, LogType.ERROR);
            }
        }
        public void Update(T obj)
        {
            try
            {
                string json = JsonConvert.SerializeObject(obj);
                Dictionary<string, object> dict = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                string set = string.Join(",", dict.Where(x => x.Key != keyName).Select(x => $"{x.Key}='{x.Value}'"));
                string sql = $"update {tableName} set {set} where {keyName} = '{dict[keyName]}'";
                DB.GetInstance().ExecuteNonQuery(sql);
            }
            catch (Exception e)
            {
                Logger.Log(e.Message, LogType.ERROR);
            }
        }

        public T GetObject(string key)
        {
            string sql = $"select * from {tableName} where {keyName} = '{key}'";
            return Query(sql);
        }
        public T GetObject(Dictionary<string, object> conditions)
        {
            string where = string.Join(" and ", conditions.Select(x => $"{x.Key}='{x.Value}'"));
            string sql = $"select * from {tableName} where {where}";
            return Query(sql);
        }

        public List<T> GetObjects(Dictionary<string, object> conditions)
        {
            string where = string.Join(" and ", conditions.Select(x => $"{x.Key}='{x.Value}'"));
            string sql = $"select * from {tableName} where {where}";
            return QueryAll(sql);
        }

        public void Update(string key, Dictionary<string, object> values)
        {
            try
            {
                string set = string.Join(",", values.Select(x => $"{x.Key}='{x.Value}'"));
                string sql = $"update {tableName} set {set} where {keyName} = '{key}'";
                DB.GetInstance().ExecuteNonQuery(sql);
            }
            catch (Exception e)
            {
                Logger.Log(e.Message, LogType.ERROR);
            }
        }


        public void Remove(Dictionary<string, object> conditions)
        {
            string where = string.Join(" and ", conditions.Select(x => $"{x.Key}='{x.Value}'"));
            string sql = $"delete from {tableName} where {where}";
            DB.GetInstance().ExecuteNonQuery(sql);
        }

        public void RemoveAll()
        {
            string sql = $"delete from {tableName}";
            DB.GetInstance().ExecuteNonQuery(sql);
        }
    }

}