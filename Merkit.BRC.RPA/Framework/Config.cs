using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Merkit.RPA.PA.Framework
{
    public static class Config
    {

        #region "Process parameters"

        public static string AppName { get; set; }
        public static bool DevelopEnvironment { get; set; }
        public static bool RunOnProduct { get; set; }
        public static int LogLevel { get; set; }
        public static string LogFileName { get; set; }
        public static string NotifyEmail { get; set; }

        public static string MsSqlHost { get; set; }
        public static string MsSqlDatabase { get; set; }        
        public static string LocalWorkFolder { get; set; }
        public static string EmailAttachmentsRootFolder { get; set; } 
        public static string ErrorExcelEmailSubject { get; set; }
        public static string ErrorExcelEmailBody { get; set; }  
        public static string MsSqlUserName { get; set; }
        public static string MsSqlPassword { get; set; }

        public static bool DebugMode { get; set; }
        public static string DebugLogFileName { get; set; }
        public static int MaxProcessableItemCount { get; set; }
        public static int ErrorWeight { get; set; }
        public static int MaxAllowedErrorCount { get; set; }

        #endregion
    }

}
