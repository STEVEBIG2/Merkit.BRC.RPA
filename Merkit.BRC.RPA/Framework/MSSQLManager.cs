using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Reflection.Emit;
using Merkit.RPA.PA.Framework;

namespace Merkit.BRC.RPA
{

    #region "MSSQL public class"

    public class MSSQLManager
    {
        public SqlConnection Connection { get; set; }
        public string ConnenctionString { get; set; }   

        #region "Private function"

        private bool MSSQLOpen()
        {
            bool needOpenClose = false;

            if (Connection == null)
            {
                Connection = new SqlConnection(ConnenctionString);
                needOpenClose = true;
            }

            if (needOpenClose)
            {
                Connection.Open();
            }

            return needOpenClose;
        }

        private void MSSQLClose(bool needOpenClose)
        {

            if (needOpenClose)
            {
                Connection.Close();
            }

        }

        #endregion

        #region "Public region"

        public MSSQLManager()
        {

        }

        public MSSQLManager(string msSqlHost, string msSqlDatabase, string userName, string password, string appCode)
        {
            ConnenctionString = MakeConnenctionString(msSqlHost, msSqlDatabase, userName, password, appCode);
        }

        public string MakeConnenctionString(string msSqlHost, string msSqlDatabase, string userName, string password, string appCode)
        {
            string conStr = String.Format("Data Source={0};Initial Catalog={1};User Id={2};Password={3};Application Name={4};Connect Timeout={5};Encrypt=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False; MultipleActiveResultSets=True", msSqlHost, msSqlDatabase, userName, password, appCode, 30);
            return conStr;
        }

        public SqlConnection Connect()
        {
            Connection = new SqlConnection(ConnenctionString);
            Connection.Open();
            return Connection;
        }

        public SqlConnection ConnectByConfig()
        {
            ConnenctionString = MakeConnenctionString(Config.MsSqlHost, Config.MsSqlDatabase, Config.MsSqlUserName, Config.MsSqlPassword, Config.AppName);
            Connection = new SqlConnection(ConnenctionString);
            return Connect();
        }

        public void Disconnect()
        {
            //con.Shutdown();
            Connection.Close();
            Connection = null;  
            //Connection.Dispose();
            return;
        }

        public object ExecuteScalar(string statement, Dictionary<string, object> arguments)
        {
            object retvalue = null;
            bool needOpenClose = MSSQLOpen();

            SqlCommand cmd = new SqlCommand(statement, Connection);

            if (arguments != null)
            {
                foreach (KeyValuePair<string, object> argument in arguments)
                {
                    //cmd.Parameters.AddWithValue("@name", "BMW");
                    cmd.Parameters.AddWithValue(argument.Key, argument.Value);
                }
            }

            retvalue = cmd.ExecuteScalar();
            cmd.Dispose();
            MSSQLClose(needOpenClose);
            return retvalue;
        }

        public DataTable ExecuteQuery(string statement, Dictionary<string, object> arguments = null, CommandType commandType = CommandType.Text, SqlTransaction tr = null)
        {
            DataTable retvalue = new DataTable();
            bool needOpenClose = MSSQLOpen();

            SqlDataAdapter sda = new SqlDataAdapter(statement, Connection);
            sda.SelectCommand.CommandType = commandType;

            if (arguments != null)
            {
                foreach (KeyValuePair<string, object> argument in arguments)
                {
                    //cmd.Parameters.AddWithValue("@name", "BMW");
                    sda.SelectCommand.Parameters.AddWithValue(argument.Key, argument.Value);
                }
            }

            if (tr != null)
            {
                sda.SelectCommand.Transaction = tr;
            }

            sda.Fill(retvalue);
            sda.Dispose();
            MSSQLClose(needOpenClose);

            return retvalue;
        }

        public int ExecuteNonQuery(string statement, Dictionary<string, object> arguments = null, SqlTransaction tr = null)
        {
            int retvalue = 0;
            bool needOpenClose = MSSQLOpen();

            SqlCommand cmd = new SqlCommand(statement, Connection);

            if (arguments != null)
            {
                foreach (KeyValuePair<string, object> argument in arguments)
                {
                    //cmd.Parameters.AddWithValue("@name", "BMW");
                    cmd.Parameters.AddWithValue(argument.Key, argument.Value);
                }
            }

            if (tr != null)
            {
                cmd.Transaction = tr;
            }

            retvalue = cmd.ExecuteNonQuery();
            cmd.Dispose();
            MSSQLClose(needOpenClose);

            return retvalue;
        }

        public int ExecuteProcWithReturnValue(string statement, Dictionary<string, object> arguments = null, SqlTransaction tr = null)
        {
            int retvalue = 0;
            bool needOpenClose = MSSQLOpen();

            SqlCommand cmd = new SqlCommand(statement, Connection);
            cmd.CommandType = CommandType.StoredProcedure;

            if (arguments != null)
            {
                foreach (KeyValuePair<string, object> argument in arguments)
                {
                    //cmd.Parameters.AddWithValue("@name", "BMW");
                    cmd.Parameters.AddWithValue(argument.Key, argument.Value);
                }
            }

            // @RETURN_VALUE
            SqlParameter returnValueParameter = new SqlParameter();
            returnValueParameter.ParameterName = "@return_value";
            returnValueParameter.SqlDbType = SqlDbType.Int;
            returnValueParameter.Direction = ParameterDirection.ReturnValue;
            cmd.Parameters.Add(returnValueParameter);

            if (tr != null)
            {
                cmd.Transaction = tr;
            }

            cmd.ExecuteNonQuery();
            //retvalue = (int)cmd.Parameters["@return_value"].Value;
            retvalue = (int)returnValueParameter.Value;

            cmd.Dispose();
            MSSQLClose(needOpenClose);

            return retvalue;
        }

        public DataTable ExecuteProcWithResults(string statement, ref int returnValue, Dictionary<string, object> arguments = null, SqlTransaction tr = null)
        {
            DataTable retvalue = new DataTable();
            bool needOpenClose = MSSQLOpen();

            SqlDataAdapter sda = new SqlDataAdapter(statement, Connection);
            sda.SelectCommand.CommandType = CommandType.StoredProcedure;

            if (arguments != null)
            {
                foreach (KeyValuePair<string, object> argument in arguments)
                {
                    //cmd.Parameters.AddWithValue("@name", "BMW");
                    sda.SelectCommand.Parameters.AddWithValue(argument.Key, argument.Value);
                }
            }

            // @RETURN_VALUE
            SqlParameter returnValueParameter = new SqlParameter();
            returnValueParameter.ParameterName = "@return_value";
            returnValueParameter.SqlDbType = SqlDbType.Int;
            returnValueParameter.Direction = ParameterDirection.ReturnValue;
            sda.SelectCommand.Parameters.Add(returnValueParameter);

            if (tr != null)
            {
                sda.SelectCommand.Transaction = tr;
            }

            sda.Fill(retvalue);
            returnValue = (int)returnValueParameter.Value;

            sda.Dispose();
            MSSQLClose(needOpenClose);

            return retvalue;
        }

        public SqlTransaction BeginTransaction()
        {

            SqlTransaction transaction = Connection.BeginTransaction();
            return transaction;
        }

        public void Commit(SqlTransaction transaction)
        {
            transaction.Commit();
            return;
        }

        public void Rollback(SqlTransaction transaction)
        {
            transaction.Rollback();
            return;
        }

        #endregion
    }

    #endregion

}
