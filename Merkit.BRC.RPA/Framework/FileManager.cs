using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Merkit.RPA.PA.Framework
{
    public static class FileManager
    {

        /// <summary>
        /// Read Text File
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static string ReadTextFile(string path)
        {
            string returnValue = File.ReadAllText(path);
            return returnValue; 
        }

        /// <summary>
        /// Logger
        /// </summary>
        /// <param name="appLogLevel"></param>
        /// <param name="currentLogLevel"></param>
        /// <param name="logFileName"></param>
        /// <param name="process"></param>
        /// <param name="logType"></param>
        /// <param name="tran"></param>
        /// <param name="note"></param>
        /// <param name="tranID"></param>
        public static void Logger(int appLogLevel, int currentLogLevel, string logFileName, string process, string logType, string tran, string note, string tranID)
        {

            // need log?
            if (currentLogLevel<=appLogLevel)
            {
                string logFileFullname = String.Format(logFileName, DateTime.Today.ToString("yyyyMMdd"));

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


        /// <summary>
        /// Get First File
        /// </summary>
        /// <param name="inputDir"></param>
        /// <param name="searchPattern"></param>
        /// <returns></returns>
        public static string GetFirstFileName(string inputDir, string searchPattern)
        {
            string workfileName = "";
            string[] files = Directory.GetFiles(inputDir, searchPattern);
            workfileName = files.FirstOrDefault();
            return workfileName;
        }

        /// <summary>
        /// Get File From Queue
        /// </summary>
        /// <param name="inputDir"></param>
        /// <param name="searchPattern"></param>
        /// <param name="workDir"></param>
        /// <returns></returns>
        public static string GetFileFromQueue(string inputDir, string searchPattern, string workDir)
        {
            string workfileName = "";
            string file = GetFirstFileName(inputDir, searchPattern);

            if (!String.IsNullOrEmpty(file))
            {
                workfileName = Path.Combine(workDir, Path.GetFileName(file)); //String.Format(@"{0}\{1}", workDir, Path.GetFileName(file));
                File.Move(file, workfileName);
            }

            return workfileName;
        }

        /// <summary>
        /// Archive Queue File
        /// </summary>
        /// <param name="queueFileName"></param>
        /// <param name="isSuccessfull"></param>
        /// <param name="successfullDir"></param>
        /// <param name="failedDir"></param>
        /// <returns></returns>
        public static string ArchiveQueueFile(string queueFileName, bool isSuccessfull, string successfullDir, string failedDir)
        {
            string workfileName = Path.Combine( isSuccessfull ? successfullDir : failedDir, Path.GetFileName(queueFileName));
            File.Move(queueFileName, workfileName);
            return workfileName;

        }

    }
}
