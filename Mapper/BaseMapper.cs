using Newtonsoft.Json;
using System;
using System.Collections.Generic;
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
        public List<T> GetAllObjects()
        {

            string sql = $"select * from '{tableName}'";
            var reader = DB.GetInstance().ExecuteReader(sql);
            List<T> objs = new List<T>();
            var columns = reader.GetSchemaTable();
            while (reader.Read())
            {
                var values = new object[reader.FieldCount];
                reader.GetValues(values);
                Dictionary<string, object> dict = new Dictionary<string, object>();
                for (int i = 0; i < columns.Rows.Count; i++)
                {
                    dict.Add(columns.Rows[i]["ColumnName"].ToString(), values[i]);
                }
                string json = JsonConvert.SerializeObject(dict);

                objs.Add(JsonConvert.DeserializeObject<T>(json));

            }
            return objs;
        }
        public void Add(T obj)
        {
            try
            {
                string json = JsonConvert.SerializeObject(obj);
                Dictionary<string, object> dict = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
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
            var reader = DB.GetInstance().ExecuteReader(sql);
            var columns = reader.GetSchemaTable();
            if (reader.Read())
            {
                var values = new object[reader.FieldCount];
                reader.GetValues(values);
                Dictionary<string, object> dict = new Dictionary<string, object>();
                for (int i = 0; i < columns.Rows.Count; i++)
                {
                    dict.Add(columns.Rows[i]["ColumnName"].ToString(), values[i]);
                }
                string json = JsonConvert.SerializeObject(dict);
                return JsonConvert.DeserializeObject<T>(json);
            }
            return null;
        }

        public List<T> GetObjects(Dictionary<string, object> conditions)
        {
            string where = string.Join(" and ", conditions.Select(x => $"{x.Key}='{x.Value}'"));
            string sql = $"select * from {tableName} where {where}";
            var reader = DB.GetInstance().ExecuteReader(sql);
            List<T> objs = new List<T>();
            var columns = reader.GetSchemaTable();
            while (reader.Read())
            {
                var values = new object[reader.FieldCount];
                reader.GetValues(values);
                Dictionary<string, object> dict = new Dictionary<string, object>();
                for (int i = 0; i < columns.Rows.Count; i++)
                {
                    dict.Add(columns.Rows[i]["ColumnName"].ToString(), values[i]);
                }
                string json = JsonConvert.SerializeObject(dict);
                objs.Add(JsonConvert.DeserializeObject<T>(json));
            }
            return objs;
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