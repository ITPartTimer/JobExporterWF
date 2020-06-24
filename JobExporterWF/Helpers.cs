using System;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;
using System.Data.Odbc;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JobExporterWF.DAL
{
    public class Helpers
    {
        private string _ODBCDataConnString;

        // --------------------------------------------
        // Database connection string
        // --------------------------------------------      

        // ODBC to Stratix using DSN names on local machine
        public string ODBCDataConnString
        {
            get
            {
                // DSN stored in System Environment variables
                string dsn = Environment.GetEnvironmentVariable("sysStratixDSN");
                string uid = Environment.GetEnvironmentVariable("sysStratixUID");
                string pwd = Environment.GetEnvironmentVariable("sysStratixPWD");

                _ODBCDataConnString = "DSN=" + dsn + ";UID=" + uid + ";Pwd=" + pwd;

                if (string.IsNullOrEmpty(_ODBCDataConnString))
                {
                    throw new NullReferenceException("Empty Connection String");
                }
                else
                {
                    return _ODBCDataConnString;
                }

            }

        }

        // Is the number integer.  True, if conversion is a success
        public static bool IsInteger (string job)
        {
            try
            {
                Convert.ToInt32(job);
                return true;
            }
            catch
            {
                return false;
            }
        }

        // ---------------------------------------------------------
        // Add input parameters to a SQL command object.  Overloaded
        // to accept a param value for input parameters
        // ---------------------------------------------------------
        protected void AddParamToSQLCmd(SqlCommand sqlCmd, string paramID,
                                        SqlDbType sqlType, int paramSize,
                                        ParameterDirection paramDirection, object paramValue)
        {
            //See VB code for some error checking to insert here

            SqlParameter newSqlParam = new SqlParameter();

            newSqlParam.ParameterName = paramID;
            newSqlParam.SqlDbType = sqlType;
            newSqlParam.Size = paramSize;
            newSqlParam.Direction = paramDirection;
            newSqlParam.Value = paramValue;

            sqlCmd.Parameters.Add(newSqlParam);
        }

        // ------------------------------------------------------
        // Overload for input parameter to allow proper config.
        // of a Decimal SqlDbType paramSize=Precision, scale=Scale
        // -------------------------------------------------------
        protected void AddParamToSQLCmd(SqlCommand sqlCmd, string paramID,
                                        SqlDbType sqlType, int paramSize, int scale,
                                        ParameterDirection paramDirection, object paramValue)
        {
            //See VB code for some error checking to insert here

            SqlParameter newSqlParam = new SqlParameter();

            newSqlParam.ParameterName = paramID;
            newSqlParam.SqlDbType = sqlType;
            newSqlParam.Precision = (byte)paramSize;
            newSqlParam.Scale = (byte)scale;
            newSqlParam.Direction = paramDirection;
            newSqlParam.Value = paramValue;

            sqlCmd.Parameters.Add(newSqlParam);
        }

        // ---------------------------------------------------
        // Add output parameters to SQL comment object
        // ---------------------------------------------------
        protected void AddParamToSQLCmd(SqlCommand sqlCmd, string paramID,
                                        SqlDbType sqlType, int paramSize,
                                        ParameterDirection paramDirection)
        {
            //See VB code for some error checking to insert here

            SqlParameter newSqlParam = new SqlParameter();

            newSqlParam.ParameterName = paramID;
            newSqlParam.SqlDbType = sqlType;
            newSqlParam.Size = paramSize;
            newSqlParam.Direction = paramDirection;

            sqlCmd.Parameters.Add(newSqlParam);
        }
    }
}
