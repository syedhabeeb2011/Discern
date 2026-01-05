using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;
using System.Reflection;
using System.Collections;
using System.Configuration;

namespace Discern_Notification
{
    /// <summary>
    /// Global Class : 
    /// Writes the all activity in a Log file and saves it in a specified location.
    /// </summary>
    static class Global
    {
        static string output;
        /// <summary>
        /// 
        /// </summary>
        static Global()
        {
            output = ConfigurationManager.AppSettings["settingsfilepath"].ToString();
        }
        /// <summary>
        /// Log : methods for performing the log write
        /// </summary>
        /// <param name="Message"></param>
        public static void WriteLog(string Message)
        {
            DateTime dateTime = System.DateTime.UtcNow.AddMinutes(330);
            if (!Directory.Exists(output))
                Directory.CreateDirectory(output);

            string logfileName = dateTime.ToString("dd_MMM_yy") + "_" + ".txt";
            StreamWriter streamwriter = new StreamWriter(output + logfileName, true);
            streamwriter.WriteLine(DateTime.Now + ": " + Message);
            streamwriter.Close();
            streamwriter.Dispose();

        }
    }
}
