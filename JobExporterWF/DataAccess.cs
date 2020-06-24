using System;
using System.ComponentModel;
using System.Data.Odbc;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JobExporterWF.Models;

namespace JobExporterWF.DAL
{
    [DataObject(true)]
    public class DataAccess : Helpers
    {
        /*
         * For testing console.writeline.   Use throw in catch blocks for production.
         * That way the caller will handle the exception.
         */
        [DataObjectMethod(DataObjectMethodType.Select)]
        public bool Find_Job (string job)
        {
            bool fnd = false;
            string j = "";

            OdbcConnection conn = new OdbcConnection(ODBCDataConnString);
            //OdbcConnection conn = new OdbcConnection("DSN=Invera;UID=livcalod;Pwd=livcalod");

            try
            {
                conn.Open();

                OdbcCommand cmd = conn.CreateCommand();
                cmd.CommandText = "select psh_job_no from iptpsh_rec where psh_job_no = " + job;
                
                //Only returning one value, so don't need a recordset
                j = cmd.ExecuteScalar().ToString();

                /*
                 * If Jog is returned, it is found.  Not found will return a null
                 * value.  Let Exceptions handle it.
                 */
                if (j.Length != 0)
                    fnd = true;

            }
            catch (OdbcException)
            {
                /*
                 * Throw exception to caller and handle it there.  If you handle
                 * here program might continue with a bad Job value causing
                 * other exceptions
                 */
                throw;
            }
            catch (Exception)
            {
                //Console.WriteLine("Find_Job other ex: " + ex.Message);
                throw;
            }
            finally
            {
                // No matter what close and dispose of the connetion
                conn.Close();
                conn.Dispose();
            }

            return fnd;

        }

        [DataObjectMethod(DataObjectMethodType.Select)]
        public List<ArborStratix> Get_Arbors (string job)
        {
            List<ArborStratix> lstArbor = new List<ArborStratix>();

            OdbcConnection conn = new OdbcConnection(ODBCDataConnString);
            //OdbcConnection conn = new OdbcConnection("DSN=Invera;UID=livcalod;Pwd=livcalod");

            try
            {
                conn.Open();

                OdbcCommand cmd = conn.CreateCommand();
                cmd.CommandText = "select * from iptfra_rec where fra_job_no = " + job;

                OdbcDataReader rdr = cmd.ExecuteReader();

                using (rdr)
                {
                    int c = 0;
                    decimal w = 0;
                    int f = rdr.FieldCount;

                    // Only one long record with 60 fra_wdth and fra_nbr_slit pairs
                    while (rdr.Read())
                    {
                        // Go thru each pair and dynamically create the field name
                        for (int i = 1; i < f; i++)
                        {
                            ArborStratix a = new ArborStratix();

                            c += 1;
                            w = (decimal)rdr["fra_wdth_" + c.ToString()];

                            // Only want a pair with a value.  When Wdth = 0, we are done
                            if (w != 0)
                            {
                                a.Job = rdr["fra_job_no"].ToString();
                                a.Wdth = (decimal)rdr["fra_wdth_" + c.ToString()];
                                a.Nbr = Convert.ToInt32(rdr["fra_nbr_slit_" + c.ToString()]);

                                lstArbor.Add(a);
                            }
                            else
                            {
                                // Jump to end of loop
                                i = f;
                            }
                                
                        }                                  
                    }
                } 
            }
            catch(OdbcException)
            {
                throw;
                //Console.WriteLine("arbor odbc ex: " + ex.Message);
            }
            catch(Exception)
            {
                throw;
                //Console.WriteLine("arbor other ex: " + ex.Message);
            }       
            finally
            {
                // No matter what close and dispose of the connetion
                conn.Close();
                conn.Dispose();
            }

            return lstArbor;
        }

        [DataObjectMethod(DataObjectMethodType.Select)]
        public List<Consume> Get_Consumed(string job)
        {
            List<Consume> lstConsume = new List<Consume>();

            OdbcConnection conn = new OdbcConnection(ODBCDataConnString);

            try
            {
                conn.Open();

                OdbcCommand cmd = conn.CreateCommand();
                cmd.CommandText = @"select frc_job_no,frc_cons_seq_no,frc_tag_no,frc_nbr_stp,frc_arb_pos,frc_cons_wgt,frc_ga_size,frc_wdth,frc_frm 
                                        from iptfrc_rec where frc_coil_no = 1 and frc_job_no = " + job + " order by frc_cons_ln_no"; 

                OdbcDataReader rdr = cmd.ExecuteReader();

                using (rdr)
                {                 
                    while (rdr.Read())
                    {
                        Consume c = new Consume();

                        c.Job = Convert.ToInt32(rdr["frc_job_no"]);                    
                        c.Seq = Convert.ToInt32(rdr["frc_cons_seq_no"]);
                        c.Tag = rdr["frc_tag_no"].ToString().Trim();                       
                        c.Stp = Convert.ToInt32(rdr["frc_nbr_stp"]);
                        c.Pos = Convert.ToInt32(rdr["frc_arb_pos"]);
                        c.Wgt = Convert.ToInt32(rdr["frc_cons_wgt"]);
                        c.Ga = (decimal)rdr["frc_ga_size"];
                        c.Wdth = (decimal)rdr["frc_wdth"];
                        c.Frm = rdr["frc_frm"].ToString();

                        lstConsume.Add(c);
                    }
                }
            }
            catch (OdbcException)
            {
                throw;
                //Console.WriteLine("consume odbc ex: " + ex.Message);
            }
            catch (Exception)
            {
                throw;
                //Console.WriteLine("consume other ex: " + ex.Message);
            }
            finally
            {
                // No matter what close and dispose of the connetion
                conn.Close();
                conn.Dispose();
            }

            return lstConsume;
        }

        [DataObjectMethod(DataObjectMethodType.Select)]
        public List<Planned> Get_Planned(string job)
        {
            List<Planned> lstPlanned = new List<Planned>();

            OdbcConnection conn = new OdbcConnection(ODBCDataConnString);

            try
            {
                conn.Open();

                OdbcCommand cmd = conn.CreateCommand();
                cmd.CommandText = @"select jpp_trgt_ord_info, left(jpp_trgt_ord_info, 2) as pfx,ltrim(substr(jpp_trgt_ord_info, 3, 8), '0') as ref, 
                                        ltrim(substr(jpp_trgt_ord_info, 11, 3), '0') as itm,right(jpp_trgt_ord_info, 2) as sitm,ppd_ga_size,ppd_wdth,
                                        jpp_part,pdt_ga_tol_posv,pdt_ga_tol_neg,pdt_wdth_tol_posv,pdt_wdth_tol_neg
                                        from iptjpp_rec p inner join cprclg_rec cps on cps.clg_part = p.jpp_part
                                        inner join cprppd_rec ppd on ppd.ppd_part_ctl_no = cps.clg_part_ctl_no
                                        inner join cprpdt_rec t on t.pdt_part_ctl_no= cps.clg_part_ctl_no
                                        where jpp_invt_typ = 'W' and jpp_part_cus_id is not null and clg_actv = 1 and jpp_job_no = " + job;

                OdbcDataReader rdr = cmd.ExecuteReader();

                using (rdr)
                {
                    while (rdr.Read())
                    {
                        Planned p = new Planned();

                        p.Trgt = rdr["jpp_trgt_ord_info"].ToString();
                        p.Pfx = rdr["pfx"].ToString();
                        p.Ref = rdr["ref"].ToString();
                        p.Itm = rdr["itm"].ToString();
                        p.SItm = rdr["sitm"].ToString();
                        p.Ga = (decimal)rdr["ppd_ga_size"];
                        p.Wdth = (decimal)rdr["ppd_wdth"];
                        p.Part = rdr["jpp_part"].ToString();
                        p.GaP = (decimal)rdr["pdt_ga_tol_posv"];
                        p.GaN = (decimal)rdr["pdt_ga_tol_neg"];
                        p.WdthP = (decimal)rdr["pdt_wdth_tol_posv"];
                        p.WdthN = (decimal)rdr["pdt_wdth_tol_neg"];

                        lstPlanned.Add(p);
                    }
                }
            }
            catch (OdbcException)
            {
                throw;
                //Console.WriteLine("consume odbc ex: " + ex.Message);
            }
            catch (Exception)
            {
                throw;
                //Console.WriteLine("consume other ex: " + ex.Message);
            }
            finally
            {
                // No matter what close and dispose of the connetion
                conn.Close();
                conn.Dispose();
            }

            //IEnumerable<Planned> lstUniquePlanned = lstPlanned.GroupBy(x => x.Part).Select(y => y.First());



            return lstPlanned.GroupBy(x => x.Part).Select(y => y.First()).ToList();

            //return lstPlanned;
        }      

        [DataObjectMethod(DataObjectMethodType.Select)]
        public Ga Get_Ga(decimal ga, List<Planned> lstPlanned)
        {
            /*
            Ga used on the Job is the Num_Size1 fomr the PPS or the Transaction Common Item.
            Use the TagNo from Planned Consumption to find the gauge in Transaction Common.

            In the second part, figure gauge range for this Job.  That will be the most constrained
            gauge range from the CPS in MultDetail.

            Finally, use Ga and Min-Max range to determine the +/- tolerance
            */
            Ga g = new Ga();

            g.Size = ga;

            decimal minGa = 0;
            decimal maxGa = 1;
            
            /*
             * Gauge min to max is the largest min and smallest max
             * of the CPS gauge tolerance listed on the Job
             */
            foreach (Planned m in lstPlanned)
            {
                // If the current is >= stored min, replace
                if ((m.Ga - m.GaN) >= minGa)
                    minGa = (m.Ga - m.GaN);

                // If the crurrent is <= stored max, replace
                if ((m.Ga + m.GaP) <= maxGa)
                    maxGa = (m.Ga + m.GaP);
            }

            // KEVIN needs a +/- tolerance, so use master coil gauge and minGa/maxGa range
            g.GaN = ga - minGa;
            g.GaP = maxGa - ga;

            return g;
        }
    }
}
