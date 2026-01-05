/*
 * Created by SharpDevelop.
 * User: v.chakradhar.ark
 * Date: 11/10/2019
 * Time: 11:00 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Text;
using System.IO;
using System.Configuration;

namespace Atos.SyntBots.Common
{
    /// <summary>
    /// Description of Logger.
    /// </summary>
    public static class Logger
    {
        private static string sLogFormat;
        private static string sErrorTime;


        public static string LogFileName
        {
            get
            {
                //this variable used to create log filename format "
                //for example filename : ErrorLogYYYYMMDDHMS
                string sYear = DateTime.Now.Year.ToString();
                string sMonth = DateTime.Now.Month.ToString();
                string sDay = DateTime.Now.Day.ToString();
                string hour = DateTime.Now.TimeOfDay.Hours.ToString();
                string minutes = DateTime.Now.TimeOfDay.Minutes.ToString();
                string second = DateTime.Now.TimeOfDay.Seconds.ToString();
                sErrorTime = "Log" + sYear + sMonth + sDay;//+ minutes + second;
                return sErrorTime;
            }
        }
        public static void Log(string message, string filepath)
        {
            FileStream fileStream = null;
            StreamWriter streamWriter = null;

            string logFilePath = filepath;

            //sLogFormat used to create log files format :
            // dd/mm/yyyy hh:mm:ss AM/PM ==> Log Message
            sLogFormat = DateTime.Now.ToShortDateString().ToString() + " " + DateTime.Now.ToString() + " ==> ";
            logFilePath = logFilePath + LogFileName + ".txt";

            #region Create the Log file directory if it does not exists

            DirectoryInfo logDirInfo = null;
            FileInfo logFileInfo = new FileInfo(logFilePath);
            logDirInfo = new DirectoryInfo(logFileInfo.DirectoryName);
            if (!logDirInfo.Exists) logDirInfo.Create();

            #endregion Create the Log file directory if it does not exists

            if (!logFileInfo.Exists)
            {
                fileStream = logFileInfo.Create();
            }
            else
            {
                fileStream = new FileStream(logFilePath, FileMode.Append);
            }
            streamWriter = new StreamWriter(fileStream);
            streamWriter.WriteLine(sLogFormat + message);

            if (streamWriter != null) streamWriter.Close();
            if (fileStream != null) fileStream.Close();
        }

        public static void Log(string message)
        {
            FileStream fileStream = null;
            StreamWriter streamWriter = null;

            string logFilePath = @"D:\TaskQueue\Logs\";

            //sLogFormat used to create log files format :
            // dd/mm/yyyy hh:mm:ss AM/PM ==> Log Message
            sLogFormat = DateTime.Now.ToShortDateString().ToString() + " " + DateTime.Now.ToString() + " ==> ";
            logFilePath = logFilePath + LogFileName + ".txt";

            #region Create the Log file directory if it does not exists

            DirectoryInfo logDirInfo = null;
            FileInfo logFileInfo = new FileInfo(logFilePath);
            logDirInfo = new DirectoryInfo(logFileInfo.DirectoryName);
            if (!logDirInfo.Exists) logDirInfo.Create();

            #endregion Create the Log file directory if it does not exists

            if (!logFileInfo.Exists)
            {
                fileStream = logFileInfo.Create();
            }
            else
            {
                fileStream = new FileStream(logFilePath, FileMode.Append);
            }
            streamWriter = new StreamWriter(fileStream);
            streamWriter.WriteLine(sLogFormat + message);

            if (streamWriter != null) streamWriter.Close();
            if (fileStream != null) fileStream.Close();
        }
    }
}
