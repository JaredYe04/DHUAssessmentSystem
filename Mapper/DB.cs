
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
using System.IO;

namespace 考核系统.Mapper
{
    internal class DB
    {
        private static string connStr = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;
        //private static SQLiteConnection cn=new SQLiteConnection(connStr);
        private static DB instance = null;


        /// <summary>
        /// 备份 SQLite 数据库到指定的文件路径
        /// </summary>
        /// <param name="sourceConnectionString">源数据库的连接字符串</param>
        /// <param name="backupFilePath">备份文件路径</param>
        public static void BackupDatabaseToSqliteFile(string backupFilePath)
        {
            try
            {
                // 打开源数据库连接
                using (var sourceConnection = new SQLiteConnection(connStr))
                {
                    sourceConnection.Open();

                    // 创建目标数据库连接
                    using (var destinationConnection = new SQLiteConnection($"Data Source={backupFilePath};Version=3;"))
                    {
                        destinationConnection.Open();

                        // 使用 BackupDatabase 方法备份数据库
                        sourceConnection.BackupDatabase(destinationConnection, "main", "main", -1, null, 0);
                        Logger.Log($"数据库成功备份到: {backupFilePath}");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"数据库备份失败: {ex.Message}",LogType.ERROR);
            }
        }

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
            // 获取程序执行文件的目录
            string exeDirectory = AppDomain.CurrentDomain.BaseDirectory;

            // 拼接数据库文件的相对路径
            string dbFilePath = Path.Combine(exeDirectory, "DhuAssessment.db");

            // 设置连接字符串（假设在配置文件中连接字符串的name是 "ConnectionString"）
            var connectionString = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;

            // 使用string.Replace替换原始的数据库路径
            connStr = connectionString.Replace("[PATH]", dbFilePath);

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

        internal static void ResetDatabase()
        {
            //清除所有表中的数据
            List<string> tables = new List<string>();
            using (SQLiteConnection cn = new SQLiteConnection(connStr))
            {
                cn.Open();
                SQLiteCommand cmd = new SQLiteCommand();
                cmd.Connection = cn;
                cmd.CommandText = "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name;";
                SQLiteDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    tables.Add(reader.GetString(0));
                }
            }
            foreach (var table in tables)
            {
                using (SQLiteConnection cn = new SQLiteConnection(connStr))
                {
                    cn.Open();
                    SQLiteCommand cmd = new SQLiteCommand();
                    cmd.Connection = cn;
                    cmd.CommandText = $"DELETE FROM {table}";
                    cmd.ExecuteNonQuery();
                    Logger.Log($"清空表{table}成功");
                }
            }
        }

        internal static void RestoreDatabaseFromSqliteFile(string fileName)
        {
            //清除原本的数据库
            ResetDatabase();
            //恢复数据库
            try
            {
                // 打开源数据库连接
                using (var sourceConnection = new SQLiteConnection($"Data Source={fileName};Version=3;"))
                {
                    sourceConnection.Open();
                    // 创建目标数据库连接
                    using (var destinationConnection = new SQLiteConnection(connStr))
                    {
                        destinationConnection.Open();
                        // 使用 BackupDatabase 方法备份数据库
                        sourceConnection.BackupDatabase(destinationConnection, "main", "main", -1, null, 0);
                        Logger.Log($"数据库成功恢复");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"数据库恢复失败: {ex.Message}", LogType.ERROR);
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
