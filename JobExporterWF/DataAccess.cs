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

            OdbcConnection conn = new OdbcConnection(STRATIXDataConnString);
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
            catch (OdbcException ex)
            {
                /*
                 * Throw exception to caller and handle it there.  If you handle
                 * here program might continue with a bad Job value causing
                 * other exceptions
                 */
                throw;
            }
            catch (Exception ex)
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

            OdbcConnection conn = new OdbcConnection(STRATIXDataConnString);
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
            catch(OdbcException ex)
            {
                throw;
                //Console.WriteLine("arbor odbc ex: " + ex.Message);
            }
            catch(Exception ex)
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

            OdbcConnection conn = new OdbcConnection(STRATIXDataConnString);

            try
            {
                conn.Open();

                OdbcCommand cmd = conn.CreateCommand();
                cmd.CommandText = "select frc_job_no,frc_cons_seq_no,frc_tag_no,frc_wdth,frc_nbr_stp,frc_arb_pos,frc_cons_wgt from iptfrc_rec where frc_coil_no = 1 and frc_job_no = " + job + " order by frc_cons_ln_no"; 

                OdbcDataReader rdr = cmd.ExecuteReader();

                using (rdr)
                {                 
                    while (rdr.Read())
                    {
                        Consume c = new Consume();

                        c.Job = Convert.ToInt32(rdr["frc_job_no"]);                    
                        c.Seq = Convert.ToInt32(rdr["frc_cons_seq_no"]);
                        c.Tag = rdr["frc_tag_no"].ToString().Trim();
                        c.Wdth = (decimal)rdr["frc_wdth"];
                        c.Stp = Convert.ToInt32(rdr["frc_nbr_stp"]);
                        c.Pos = Convert.ToInt32(rdr["frc_arb_pos"]);
                        c.Wgt = Convert.ToInt32(rdr["frc_cons_wgt"]);

                        lstConsume.Add(c);
                    }
                }
            }
            catch (OdbcException ex)
            {
                throw;
                //Console.WriteLine("consume odbc ex: " + ex.Message);
            }
            catch (Exception ex)
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
        public List<SO> Get_SOs(string job)
        {
            List<SO> lstSO = new List<SO>();

            OdbcConnection conn = new OdbcConnection(STRATIXDataConnString);

            try
            {
                conn.Open();

                // Try to split with verbatim literal
                OdbcCommand cmd = conn.CreateCommand();
                cmd.CommandText = @"select jpp_trgt_ord_info, left(jpp_trgt_ord_info, 2) as pfx, ltrim(substr(jpp_trgt_ord_info, 3, 8), '0') as ref, 
                                        ltrim(substr(jpp_trgt_ord_info, 11, 3), '0') as itm, 
                                        right(jpp_trgt_ord_info, 2) as sitm from iptjpp_rec where jpp_invt_typ = 'W' and jpp_part_cus_id is not null and jpp_job_no = " + job;

                OdbcDataReader rdr = cmd.ExecuteReader();

                using (rdr)
                {
                    while (rdr.Read())
                    {
                        SO s = new SO();

                        s.Trgt = rdr["jpp_trgt_ord_info"].ToString();
                        s.Pfx = rdr["pfx"].ToString();
                        s.Ref = rdr["ref"].ToString();
                        s.Itm = rdr["itm"].ToString();
                        s.SItm = rdr["sitm"].ToString();

                        lstSO.Add(s);
                    }
                }
            }
            catch (OdbcException ex)
            {
                throw;
                //Console.WriteLine("SO odbc ex: " + ex.Message);
            }
            catch (Exception ex)
            {
                throw;
                //Console.WriteLine("SO other ex: " + ex.Message);
            }
            finally
            {
                // No matter what close and dispose of the connetion
                conn.Close();
                conn.Dispose();
            }

            return lstSO;
        }

        [DataObjectMethod(DataObjectMethodType.Select)]
        public List<MultDetail> Get_Details(string qry)
        {
            List<MultDetail> lstDetail = new List<MultDetail>();

            OdbcConnection conn = new OdbcConnection(STRATIXDataConnString);

            try
            {
                conn.Open();

                // Try to split with verbatim literal
                OdbcCommand cmd = conn.CreateCommand();
                cmd.CommandText = @"select ipd_ref_pfx, ipd_ref_no, ipd_ref_itm, ipd_cus_ven_id, ipd_part, ipd_part_ctl_no,ipd_frm, ipd_ga_size, 
                                        pdt_ga_tol_posv,pdt_ga_tol_neg,ipd_wdth, pdt_wdth_tol_posv,pdt_wdth_tol_neg
                                        from tctipd_rec ipd inner join cprpdt_rec pdt on ipd.ipd_part_ctl_no = pdt.pdt_part_ctl_no 
                                        where ipd_ref_no || ipd_ref_itm in (" + qry + ") and ipd_ref_pfx = 'SO'";

                OdbcDataReader rdr = cmd.ExecuteReader();

                using (rdr)
                {
                    while (rdr.Read())
                    {
                        MultDetail m = new MultDetail();                     
                        m.Pfx = rdr["ipd_ref_pfx"].ToString();
                        m.Ref = rdr["ipd_ref_no"].ToString();
                        m.Itm = rdr["ipd_ref_itm"].ToString();
                        m.Cus = rdr["ipd_cus_ven_id"].ToString().Trim();
                        m.Part = rdr["ipd_part"].ToString().Trim();
                        m.CtlNo = Convert.ToInt32(rdr["ipd_part_ctl_no"]);
                        m.Frm = rdr["ipd_frm"].ToString().Trim();
                        m.Ga = (decimal)rdr["ipd_ga_size"];
                        m.GaP = (decimal)rdr["pdt_ga_tol_posv"];
                        m.GaN = (decimal)rdr["pdt_ga_tol_neg"];
                        m.Wdth = (decimal)rdr["ipd_wdth"];
                        m.WdthP = (decimal)rdr["pdt_wdth_tol_posv"];
                        m.WdthN = (decimal)rdr["pdt_wdth_tol_neg"];

                        lstDetail.Add(m);
                    }
                }
            }
            catch (OdbcException ex)
            {
                throw;
                //Console.WriteLine("MultDetail odbc ex: " + ex.Message);
            }
            catch (Exception ex)
            {
                throw;
                //Console.WriteLine("MultDetail other ex: " + ex.Message);
            }
            finally
            {
                // No matter what close and dispose of the connetion
                conn.Close();
                conn.Dispose();
            }

            return lstDetail;
        }

        [DataObjectMethod(DataObjectMethodType.Select)]
        public Ga Get_Ga(string tag, List<MultDetail> lstDetail)
        {
            /*
            Ga used on the Job is the Num_Size1 fomr the PPS or the Transaction Common Item.
            Use the TagNo from Planned Consumption to find the gauge in Transaction Common.

            In the second part, figure gauge range for this Job.  That will be the most constrained
            gauge range from the CPS in MultDetail.

            Finally, use Ga and Min-Max range to determine the +/- tolerance
            */
            Ga g = new Ga();

            OdbcConnection conn = new OdbcConnection(STRATIXDataConnString);

            try
            {
                conn.Open();

                OdbcCommand cmd = conn.CreateCommand();
                cmd.CommandText = "select ipd_num_size1 "
                                        + "from iptfrc_rec frc inner join intpcr_rec pcr on frc.frc_itm_ctl_no = pcr.pcr_itm_ctl_no "
                                        + "inner join tctipd_rec ipd on pcr.pcr_po_pfx = ipd.ipd_ref_pfx "
                                        + "and pcr.pcr_po_no = ipd.ipd_ref_no and pcr.pcr_po_itm = ipd.ipd_ref_itm "
                                        + "inner join pnttol_rec tol on pcr.pcr_po_pfx = tol.tol_ref_pfx "
                                        + "and pcr.pcr_po_no = tol.tol_ref_no and pcr.pcr_po_itm = tol.tol_ref_itm "
                                        + "where frc.frc_tag_no = " + "'" + tag.ToString() + "'";

                // Only returning one value, so don't need a recordset
                g.NumSize =  Convert.ToDecimal(cmd.ExecuteScalar());


               
            }
            catch (OdbcException ex)
            {
                throw;
                //Console.WriteLine("Ga odbc ex: " + ex.Message);
            }
            catch (Exception ex)
            {
                throw;
                //Console.WriteLine("Ga other ex: " + ex.Message);
            }
            finally
            {
                // No matter what close and dispose of the connetion
                conn.Close();
                conn.Dispose();
            }

            decimal minGa = 0;
            decimal maxGa = 1;
            
            /*
             * Gauge min to max is the largest min and smallest max
             * of the CPS gauge tolerance listed on the Job
             */
            foreach (MultDetail m in lstDetail)
            {
                // If the current is >= stored min, replace
                if ((m.Ga - m.GaN) >= minGa)
                    minGa = (m.Ga - m.GaN);

                // If the crurrent is <= stored max, replace
                if ((m.Ga + m.GaP) <= maxGa)
                    maxGa = (m.Ga + m.GaP);
            }

            // KEVIN needs a +/- tolerance, so use master coil gauge and minGa/maxGa range
            g.GaN = g.NumSize - minGa;
            g.GaP = maxGa - g.NumSize;

            return g;
        }
    }
}
