using System;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Reflection;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JobExporterWF.Log
{
    public class Logger
    {
        /*
        * How to use Logger.  Capable of writing Exception or string to either File or Db
        * 1. Add existing class to project: Logger.cs
        * 2. Add to app.config
        * <add key="LogTyp" value="FL"/>
        * <add key="LogPath" value="C:\ConsoleAppsToRun\Log.txt"/>
        * 3. Logger.LogWrite(<Type>,<DB> or <FL>, <Exception> or <String message>)
        */

        protected const int _max50 = 50;
        protected const int _max250 = 250;

        protected string logHow { get; set; }
        protected string logPath { get; set; }

        #region Properties
        private string _Typ = string.Empty;
        private string _App = string.Empty;
        private string _Msg = string.Empty;
        private string _Src = "NA";
        private string _Trc = "NA";
        private string _Trgt = "NA";

        public string Typ
        {
            get
            {
                return _Typ;
            }
            set
            {
                _Typ = value;
            }
        }

        public string App
        {
            get
            {
                return _App;
            }
            set
            {
                _App = value;
            }
        }

        public string Msg
        {
            get
            {
                return _Msg;
            }
            set
            {
                if (value.Length > _max250)
                    _Msg = value.Substring(0, _max250 - 1);
                else
                    _Msg = value;
            }
        }

        public string Src
        {
            get
            {
                return _Src;
            }
            set
            {
                if (value.Length > _max50)
                    _Src = value.Substring(0, _max50 - 1);
                else
                    _Src = value;
            }
        }

        public string Trc
        {
            get
            {
                return _Trc;
            }
            set
            {
                if (value.Length > _max50)
                    _Trc = value.Substring(0, _max50 - 1);
                else
                    _Trc = value;
            }
        }

        public string Trgt
        {
            get
            {
                return _Trgt;
            }
            set
            {
                if (value.Length > _max50)
                    _Trgt = value.Substring(0, _max50 - 1);
                else
                    _Trgt = value;
            }
        }

        #endregion

        /*
         * Constructor gets the configuration settings.
         * Config file must be located in same folder as executable program
         */
        public Logger()
        {
            logHow =  ConfigurationManager.AppSettings.Get("LogHow");
            logPath = ConfigurationManager.AppSettings.Get("LogPath");           
        }


        /*
         * LogWrite is overloaded to write an Exception
         * or just a message from the program
         */
        public static void LogWrite(string typ, Exception ex)
        {
            Logger l = new Logger();

            l.Typ = typ;
            l.App = AppDomain.CurrentDomain.FriendlyName;
            l.Msg = ex.Message.ToString();
            l.Src = ex.Source.ToString();
            l.Trc = ex.StackTrace.ToString();
            l.Trgt = ex.TargetSite.ToString();

            if (l.logHow == "DB")
                WriteDB(l);
            else
                WriteFile(l);
        }

        public static void LogWrite(string typ, string msg)
        {
            Logger l = new Logger();

            l.Typ = typ;
            l.App = AppDomain.CurrentDomain.FriendlyName;
            l.Msg = msg;        

            if (l.logHow == "DB")
                WriteDB(l);
            else
                WriteFile(l);
        }

        // Method internal to Logger classs
        protected static void WriteDB (Logger l)
        {
            SqlConnection conn = new SqlConnection("Data Source=VC-030\\SQLDEV2014VC030;Initial Catalog=LoggerData;Integrated Security=SSPI");

            try
            {
                conn.Open();

                SqlCommand cmd = new SqlCommand();

                cmd.CommandType = CommandType.Text;
                cmd.Connection = conn;

                cmd.CommandText = "INSERT INTO Log (Typ,AppName,Msg,Src,StkTrc,Trgt) VALUES (@typ,@app,@msg,@src,@trc,@trgt)";

                cmd.Parameters.Add("@typ", SqlDbType.VarChar);
                cmd.Parameters.Add("@app", SqlDbType.VarChar);
                cmd.Parameters.Add("@msg", SqlDbType.VarChar);
                cmd.Parameters.Add("@src", SqlDbType.VarChar);
                cmd.Parameters.Add("@trc", SqlDbType.VarChar);
                cmd.Parameters.Add("@trgt", SqlDbType.VarChar);

                cmd.Parameters[0].Value = l.Typ;
                cmd.Parameters[1].Value = l.App;
                cmd.Parameters[2].Value = l.Msg;
                cmd.Parameters[3].Value = l.Src;
                cmd.Parameters[4].Value = l.Trc;
                cmd.Parameters[5].Value = l.Trgt;             

                cmd.ExecuteNonQuery();

            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                // No matter what close and dispose of the connetion
                conn.Close();
                conn.Dispose();
            }
        }

        protected static void WriteFile (Logger l)
        {
            // Get Log location from app.config
            string logPath = ConfigurationManager.AppSettings.Get("LogPath");

            try
            {
                using (StreamWriter writer = new StreamWriter(l.logPath, true))
                {
                    writer.WriteLine("----------------------------");
                    writer.WriteLine("DT = " + DateTime.Now.ToString());

                    PropertyInfo[] props = l.GetType().GetProperties();
                    foreach (PropertyInfo p in props)
                    {
                        writer.WriteLine(p.Name + " = " + p.GetValue(l, null).ToString());
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }           
        }
    }
}
