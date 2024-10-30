using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 考核系统.Utils
{
    public enum LogType
    {
        INFO,
        WARNING,
        ERROR
    }
    public class Logger
    {
        public static TextBox logger { set; get; } = null;

        public static void Log(string message, LogType type=LogType.INFO)
        {
            if (logger == null)
            {
                return;
            }
            logger.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss "));
            switch (type)
            {
                case LogType.INFO:
                    logger.AppendText("[INFO] ");
                    break;
                case LogType.WARNING:
                    logger.AppendText("[WARNING] ");
                    break;
                case LogType.ERROR:
                    logger.AppendText("[ERROR] " );
                    break;
            }
            logger.AppendText(message + "\r\n");
        }
    }
}
