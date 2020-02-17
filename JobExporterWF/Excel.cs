using System;
using System.IO;
using System.ComponentModel;
using System.Configuration;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JobExporterWF.Models;

namespace JobExporterWF.XLS
{
    [DataObject(true)]
    public class ExcelExport
    {
        public void WriteHdr(List<HdrFile> hdr)
        {
            // Copy an empty file to the destination each time
            string fileName = ConfigurationManager.AppSettings.Get("HdrFileName");
            string templatePath = ConfigurationManager.AppSettings.Get("TemplatePath");
            string destPath = ConfigurationManager.AppSettings.Get("DestPath");

            File.Copy(Path.Combine(templatePath, fileName), Path.Combine(destPath, fileName), true);

            // initialize text used in OleDbCommand
            string cmdText = "";

            string excelConnString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path.Combine(destPath, fileName) + @";Extended Properties=""Excel 8.0;HDR=YES;""";

            using (OleDbConnection eConn = new OleDbConnection(excelConnString))
            {
                try
                {
                    eConn.Open();

                    OleDbCommand eCmd = new OleDbCommand();

                    eCmd.Connection = eConn;

                    // Loop through each record and add yesterday details to XLS
                    foreach (HdrFile m in hdr)
                    {
                        // Use parameters to insert into XLS
                        cmdText = "Insert into [Sheet1$] (JobRef,Customer,Material,Width,Thickness,KnifeClear,[Clear],GaugePlus,GaugeMinus,[Note],Weight,ArbName,Reg,KnifeSet,Sep,SepClr,RingSet)" + 
                                            "Values(@job,@cust,@mtl,@wdth,@thk,@kclear,@clear,@gplus,@gminus,@note,@wgt,@arb,@reg,@kset,@sep,@sepclr,@rset)";

                        eCmd.CommandText = cmdText;

                        eCmd.Parameters.AddRange(new OleDbParameter[]
                        {
                                    new OleDbParameter("@job", m.Job),
                                    new OleDbParameter("@cust", m.Cust),
                                    new OleDbParameter("@mtl", m.Mtl),
                                    new OleDbParameter("@wdth", m.Wdth.ToString()),
                                    new OleDbParameter("@thk", m.Ga.ToString()),
                                    new OleDbParameter("@kclear", m.KnifeClr.ToString()),
                                    new OleDbParameter("@clear", m.Clr.ToString()),
                                    new OleDbParameter("@gplus", m.GaP.ToString()),
                                    new OleDbParameter("@gminus", m.GaN.ToString()),
                                    new OleDbParameter("@note", m.Note),
                                    new OleDbParameter("@wgt", m.Wgt.ToString()),
                                    new OleDbParameter("@arb", m.ArbName),
                                    new OleDbParameter("@reg", m.Reg),
                                    new OleDbParameter("@kset", m.KnifeSet),
                                    new OleDbParameter("@sep", m.Sep.ToString()),
                                    new OleDbParameter("@sepclr", m.SepClr.ToString()),
                                    new OleDbParameter("@rset", m.RingSet.ToString()),
                        });

                        eCmd.ExecuteNonQuery();

                        // Need to clear Parameters on each pass
                        eCmd.Parameters.Clear();
                    }
                }
                catch (OleDbException ex)
                {
                    throw;
                    //Console.WriteLine("OleDb Hdr error: " + ex.Message);
                }
            }
        }

        public void WriteMults(List<MultFile> mults)
        {
            // Copy an empty file to the destination each time
            string fileName = ConfigurationManager.AppSettings.Get("MultFileName");
            string templatePath = ConfigurationManager.AppSettings.Get("TemplatePath");
            string destPath = ConfigurationManager.AppSettings.Get("DestPath");

            File.Copy(Path.Combine(templatePath, fileName), Path.Combine(destPath, fileName), true);

            // initialize text used in OleDbCommand
            string cmdText = "";

            string excelConnString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path.Combine(destPath, fileName) + @";Extended Properties=""Excel 8.0;HDR=YES;""";

            using (OleDbConnection eConn = new OleDbConnection(excelConnString))
            {
                try
                {
                    eConn.Open();

                    OleDbCommand eCmd = new OleDbCommand();

                    eCmd.Connection = eConn;

                    // Loop through each record and add yesterday details to XLS
                    foreach (MultFile m in mults)
                    {
                        // Use parameters to insert into XLS
                        cmdText = "Insert into [Sheet1$] (JobRef,Customer,Quantity,[Size],WidthPlus,WidthMinus,KnifeSize)" +
                            "Values(@job,@cust,@qty,@wdth,@wplus,@wminus,@knife)";

                        eCmd.CommandText = cmdText;

                        eCmd.Parameters.AddRange(new OleDbParameter[]
                        {
                                    new OleDbParameter("@job", m.Job),
                                    new OleDbParameter("@cust", m.Cust),
                                    new OleDbParameter("@qty", m.Qty.ToString()),
                                    new OleDbParameter("@wdth", m.Size.ToString()),
                                    new OleDbParameter("@wplus", m.WdthP.ToString()),
                                    new OleDbParameter("@wminus", m.WdthN.ToString()),
                                    new OleDbParameter("@knife", m.Knife.ToString()),                            
                        });

                        eCmd.ExecuteNonQuery();

                        // Need to clear Parameters on each pass
                        eCmd.Parameters.Clear();
                    }
                }
                catch (Exception ex)
                {
                    throw;
                    //Console.WriteLine("OleDb Mult error: " + ex.Message);
                }
            }
        }

        
    }
}
