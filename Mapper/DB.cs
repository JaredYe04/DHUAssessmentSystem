
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


        private static DB instance = null;
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
            SQLiteConnection cn = new SQLiteConnection(connStr);
            cn.Open();//打开数据库
            SQLiteCommand cmd = new SQLiteCommand();
            cmd.Connection = cn;//把 SQLiteCommand的 Connection和SQLiteConnection 联系起来
            cmd.CommandText = sql;//输入SQL语句
            cmd.ExecuteNonQuery();//调用此方法运行
            cn.Close();//关闭数据库
        }
        public SQLiteDataReader ExecuteReader(string sql)
        {
            SQLiteConnection cn = new SQLiteConnection(connStr);
            cn.Open();//打开数据库
            SQLiteCommand cmd = new SQLiteCommand();
            cmd.Connection = cn;//把 SQLiteCommand的 Connection和SQLiteConnection 联系起来
            cmd.CommandText = sql;//输入SQL语句
            SQLiteDataReader reader = cmd.ExecuteReader();
            return reader;
        }
        public void ExecuteNonQuery(string sql, Dictionary<string, object> parameters)
        {
            SQLiteConnection cn = new SQLiteConnection(connStr);
            cn.Open();//打开数据库
            SQLiteCommand cmd = new SQLiteCommand();
            cmd.Connection = cn;//把 SQLiteCommand的 Connection和SQLiteConnection 联系起来
            cmd.CommandText = sql;//输入SQL语句
            foreach (var item in parameters)
            {
                cmd.Parameters.AddWithValue(item.Key, item.Value);
            }
            cmd.ExecuteNonQuery();//调用此方法运行
            cn.Close();//关闭数据库
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
