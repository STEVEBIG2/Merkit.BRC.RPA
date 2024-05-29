using CredentialManagement;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Merkit.RPA.PA.Framework;
using System.Text.RegularExpressions;

namespace Merkit.BRC.RPA
{
    public static class Framework
    {

        /// <summary>
        /// Version
        /// </summary>
        /// <returns></returns>
        public static string VersionInfo()
        {
            return "0.5.29";
        }

        /// <summary>
        /// Is URl valid?
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public static bool IsValidURL(string url)
        {
            string Pattern = @"^(?:http(s)?:\/\/)?[\w.-]+(?:\.[\w\.-]+)+[\w\-\._~:/?#[\]@!\$&'\(\)\*\+,;=.]+$";
            Regex Rgx = new Regex(Pattern, RegexOptions.Compiled | RegexOptions.IgnoreCase);
            bool isOk = Rgx.IsMatch(url);
            return isOk;
        }

        /// <summary>
        /// Is URl valid?
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public static bool IsValidZip(string url)
        {
            string Pattern = @"^(?:http(s)?:\/\/)?[\w.-]+(?:\.[\w\.-]+)+[\w\-\._~:/?#[\]@!\$&'\(\)\*\+,;=.]+$";
            Regex Rgx = new Regex(Pattern, RegexOptions.Compiled | RegexOptions.IgnoreCase);
            bool isOk = Rgx.IsMatch(url);
            return isOk;
        }

        /// <summary>
        /// Logger
        /// </summary>
        /// <param name="currentLogLevel"></param>
        /// <param name="process"></param>
        /// <param name="logType"></param>
        /// <param name="tran"></param>
        /// <param name="tranID"></param>
        /// <param name="note"></param>
        public static void Logger(int currentLogLevel, string process, string logType, string tran, string tranID, string note)
        {
            // need log?
            if (currentLogLevel <= Config.LogLevel)
            {
                string logFileFullname = String.Format(Config.LogFileName, DateTime.Today.ToString("yyyyMMdd"));
                
                // not exists log file?
                if (!File.Exists(logFileFullname))
                {
                    // create new log file
                    using (StreamWriter sw = File.CreateText(logFileFullname))
                    {
                        sw.WriteLine("Date,Time,Process,LogType,Tran,Item,Note");
                    }
                }
                
                // append to log file
                using (StreamWriter sw = File.AppendText(logFileFullname))
                {
                    sw.WriteLine(String.Format(DateTime.Now.ToString("yyyy-MM-dd") + "," + DateTime.Now.ToString("HH:mm:ss") + ",{0},{1},{2},{3},{4}", process, logType, tran, tranID, note));
                }
                
            }

            return;
        }
    }

    /// <summary>
    /// Password Repository
    /// </summary>
    public static class PasswordRepository
    {

        /// <summary>
        /// Save PasswordSaveWindows Credential
        /// </summary>
        /// <param name="passwordName"></param>
        /// <param name="userName"></param>
        /// <param name="password"></param>
        public static void SaveWindowsCredential(string passwordName, string userName, string password)
        {
            using (var cred = new Credential())
            {
                cred.Username = userName;
                cred.Password = password;
                cred.Target = passwordName;
                cred.Type = CredentialType.Generic;
                cred.PersistanceType = PersistanceType.LocalComputer;
                cred.Save();
            }
        }

        /// <summary>
        /// Get Windows Credentil
        /// </summary>
        /// <param name="passwordName"></param>
        /// <param name="userName"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public static bool GetWindowsCredential(string passwordName, ref string userName, ref string password)
        {
            using (var cred = new Credential())
            {
                cred.Target = passwordName;
                cred.Load();
                userName = String.IsNullOrEmpty(cred.Username) ? "": cred.Username;   
                password = String.IsNullOrEmpty(cred.Password) ? "" : cred.Password;
                return !String.IsNullOrEmpty(userName);
            }
        }

        /// <summary>
        /// Get Windows Credentil
        /// </summary>
        /// <param name="passwordName"></param>
        /// <param name="userName"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public static void DeleteWindowsCredential(string passwordName)
        {
            using (var cred = new Credential())
            {
                cred.Target = passwordName;
                cred.Delete();
                return;
            }
        }
    }
}
