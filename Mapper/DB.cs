
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using 考核系统.Entity;
using 考核系统.Utils;
using System.Configuration;

using System.Data.SQLite;

namespace 考核系统.Mapper
{
    internal class DB
    {
        private static string connStr = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;
        //private static SQLiteConnection cn=new SQLiteConnection(connStr);
        private static DB instance = null;
        //public static void Open()
        //{
        //    if(cn.State != System.Data.ConnectionState.Open)
        //        cn.Open();
        //}
        //public static void Close()
        //{
        //    if (cn.State != System.Data.ConnectionState.Closed)
        //        cn.Close();
        //}
        public static DB GetInstance()
        {
            //加一个互斥锁，防止多线程同时访问
            if(instance == null)
            {
                instance = new DB();
            }
            return instance;
        }
        private DB()
        {
            Logger.Log("数据库连接成功");
        }
        public void ExecuteNonQuery(string sql)
        {
            using (SQLiteConnection cn = new SQLiteConnection(connStr))
            {
                cn.Open();
                SQLiteCommand cmd = new SQLiteCommand();
                cmd.Connection = cn;
                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();
            }
        }
        //public SQLiteDataReader ExecuteReader(string sql, out SQLiteConnection connection)
        //{
        //    SQLiteConnection cn = new SQLiteConnection(connStr);

        //    cn.Open();
        //    SQLiteCommand cmd = new SQLiteCommand();
        //    cmd.Connection = cn;
        //    cmd.CommandText = sql;
        //    SQLiteDataReader reader = cmd.ExecuteReader();
        //    connection = cn;//让用户自己关闭连接
        //    return reader;

        //}

        ////重写ExecuteReader方法，直接返回读取到的数据

        public List<Dictionary<string, object>> ExecuteReader(string sql)
        {
            using (SQLiteConnection cn = new SQLiteConnection(connStr))
            {
                cn.Open();
                SQLiteCommand cmd = new SQLiteCommand();
                cmd.Connection = cn;
                cmd.CommandText = sql;
                SQLiteDataReader reader = cmd.ExecuteReader();
                List<Dictionary<string, object>> result = new List<Dictionary<string, object>>();
                while (reader.Read())
                {
                    Dictionary<string, object> dict = new Dictionary<string, object>();
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        dict.Add(reader.GetName(i), reader.GetValue(i));
                    }
                    result.Add(dict);
                }
                return result;
            }
        }

        public void ExecuteNonQuery(string sql, Dictionary<string, object> parameters)
        {
            using (SQLiteConnection cn = new SQLiteConnection(connStr))
            {
                cn.Open();
                SQLiteCommand cmd = new SQLiteCommand();
                cmd.Connection = cn;
                cmd.CommandText = sql;
                foreach (var item in parameters)
                {
                    cmd.Parameters.AddWithValue(item.Key, item.Value);
                }
                cmd.ExecuteNonQuery();
            }
        }
        ~DB()
        {

        }
    }
    static class TestDB
    {
        static void Main()
        {
            var managerMapper = ManagerMapper.GetInstance();
            managerMapper.Add(new Manager(1, "m1", "tom"));
            managerMapper.Add(new Manager(2, "m2", "jerry"));

            var list= managerMapper.GetAllObjects();


        }
    }
}
