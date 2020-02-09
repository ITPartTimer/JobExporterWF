using System;
using System.IO;
using System.Collections.Generic;
using System.Configuration;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using JobExporterWF.Models;
using JobExporterWF.DAL;
using JobExporterWF.XLS;

namespace JobExporterWF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Init erro and progress bar
            lblError.Text = "";
            lblFiles.Text = "";
            pBar.Value = 0;
        }

        // This method will not execute unless txtJob passes validation
        private void btnExport_Click(object sender, EventArgs e)
        {         
            string job = txtJob.Text;

            // Clear
            lblError.Text = "";
            lblFiles.Text = "";
            pBar.Value = 0;
            lvHeader.Clear();
            lvMults.Clear();

            /*
             * Find Job in schedule iptpsh_rec.  If not found, a null 
             * value is returned with a generic exception message. Format
             * error message and return on catch so program does not continue.
             */
            DataAccess objFindJob = new DataAccess();

            bool fnd = false;

            try
            {
                fnd = objFindJob.Find_Job(job);
            }
            catch (Exception ex)
            {
                string msg = ex.Message;

                if (msg.Substring(0, 6) == "Object")
                    msg = "Job not found in schedule";

                lblError.Text = msg;

                return;
            }
           
            #region Consume
            /*
            Get Planned Consumption from Stratix.
            The number of rows is the number of setups on this Job
            Sort by Seq.  Use Pos information to determine where 
            at what arbor position to start each setup
            */
            DataAccess objConsume = new DataAccess();

            List<Consume> lstConsume = new List<Consume>();

            try
            {
                lstConsume = objConsume.Get_Consumed(job);
            }
            catch (Exception ex)
            {
                lblError.Text = ex.Message;
                return;
            }
                     
            // test
            //Console.WriteLine("==== CONSUME ====");
            //foreach (Consume c in lstConsume)
            //{
            //    Console.WriteLine(c.Job.ToString() + " / " + c.Seq.ToString() + " / " + c.Tag + " / " + c.Wdth.ToString() + " / " + c.Stp.ToString() + " / " + c.Pos.ToString());
            //}

            pBar.Value = 12;
            #endregion           

            #region SO
            /*
            Get Unique SOs from Stratix.
            Planned Production iptjpp_rec will list each SO on the Job
            Parse this into a Prefix, Ref, Item and SubItem.
            This will be used to query Transaction Common Items joined with
            CPS Tolerances to get item details.           
            */
            DataAccess objSO = new DataAccess();

            List<SO> lstSO = new List<SO>();

            try
            {
                lstSO = objSO.Get_SOs(job);
            }
            catch (Exception ex)
            {
                lblError.Text = ex.Message;
                return;
            }
            
            // test
            //Console.WriteLine("==== SO ====");
            //foreach (SO s in lstSO)
            //{
            //    Console.WriteLine(s.Trgt + " / " + s.Pfx + " / " + s.Ref + " / " + s.Itm + " / " + s.SItm);
            //}

            // Build a list of Ref + Item to be used in the query for Item Details
            string inQry = "";

            foreach (SO s in lstSO)
            {
                inQry = inQry + string.Concat(s.Ref, s.Itm, ",");
            }

            inQry = inQry.TrimEnd(',');

            pBar.Value = 24;

            // test
            //Console.WriteLine("inQry: " + inQry);
            #endregion

            #region MultDetail
            /*
            Transaction Common Item Product joined with CPS Tolerances will
            give you all the detail you need about the cut for each SO (ex:56572-54)
            on the Job
            */
            DataAccess objDetail = new DataAccess();

            List<MultDetail> lstDetail = new List<MultDetail>();

            try
            {
                lstDetail = objDetail.Get_Details(inQry);
            }
            catch (Exception ex)
            {
                lblError.Text = ex.Message;
                return;
            }

            // test
            //Console.WriteLine("==== MULT DETAIL ====");
            //foreach (MultDetail m in lstDetail)
            //{
            //    Console.WriteLine(m.Pfx + " / " + m.Ref + " / " + m.Itm + " / " + m.Cus + " / " + m.Part + " / " + m.CtlNo.ToString() + " / " + m.Frm + " / " + m.Ga.ToString() + " / " + m.GaP.ToString() + " / " + m.GaN.ToString() + " / " + m.Wdth.ToString() + " / " + m.WdthP.ToString() + " / " + m.WdthN.ToString());
            //}

            pBar.Value = 36;
            #endregion

            #region Ga
            /*
            Get the Gauge listed on the Job
            
            Determine the most constrained gauge range from the CPS on the Job
            */

            // Consume query was ordered by seq, so first member contains the TagNo for the job
            string tag = lstConsume.Select(x => x.Tag).First().ToString();

            DataAccess objGa = new DataAccess();

            Ga g = new Ga();

            try
            {
                g = objGa.Get_Ga(tag, lstDetail);
            }
            catch (Exception ex)
            {
                lblError.Text = ex.Message;
                return;
            }

            pBar.Value = 48;
            #endregion

            #region Build HdrFile          
            /*
            Build List<HdrFile>
            */

            // Determine number of setups
            int numSetups = lstConsume.Count;
            List<string> lstNumSetups = new List<string>();

            //Add sufix of 1,2,3... to end of Job
            if (numSetups == 1)
                lstNumSetups.Add(string.Concat(job.ToString(), "-", "1"));
            else
                for (int i = 0; i < numSetups; i++)
                    lstNumSetups.Add(string.Concat(job.ToString(), "-", (i + 1).ToString()));


            List<HdrFile> lstHdr = new List<HdrFile>();

            foreach (string j in lstNumSetups)
            {
                HdrFile h = new HdrFile();

                h.Job = j;
                h.Mtl = lstDetail[0].Frm;
                h.Wdth = lstConsume[0].Wdth;
                h.Ga = g.NumSize;
                h.Clr = h.KnifeClr * h.Ga;
                h.GaP = g.GaP;
                h.GaN = g.GaN;

                // After Pos = 1, if consecutive Pos are same value, you are reslitting
                //FIX THIS TO ADD RESLIT TO NOTE
                //if (lstConsume[j].Pos != 1)
                //{
                //    if (lstConsume[j].Pos == lstConsume[j - 1].Pos)
                //        h.Note = "RESLIT";
                //}

                lstHdr.Add(h);
            }
          
            pBar.Value = 60;

            //testing
            Console.WriteLine("==== HDR FILE ====");
            foreach (HdrFile h in lstHdr)
            {
                Console.WriteLine(h.Job + " / " + h.Cust + " / " + h.Mtl + " / " + h.Wdth.ToString() + " / " + h.Ga.ToString() + " / " + h.KnifeClr.ToString() + " / " + h.Clr.ToString() + " / " + h.GaP.ToString() + " / " + h.GaN.ToString() + " / " + h.Note);
            }
            #endregion

            #region Build MultFile
            /*
            Build MultFile
            */

            // Get Arbor from Stratix
            DataAccess objArbor = new DataAccess();

            List<ArborStratix> lstArborStratix = new List<ArborStratix>();

            try
            {
                lstArborStratix = objArbor.Get_Arbors(job);
            }
            catch (Exception ex)
            {
                lblError.Text = ex.Message;
                return;
            }

            // test
            //Console.WriteLine("==== ARBOR ====");
            //foreach (ArborStratix a in lstArborStratix)
            //{
            //    Console.WriteLine(a.Job.ToString() + " / " + a.Wdth.ToString() + " / " + a.Nbr.ToString());
            //}

            // Get list of start Pos for each setup
            List<int> lstStartPos = new List<int>();

            foreach (Consume c in lstConsume)
                lstStartPos.Add(c.Pos);

            // Expand lstArborStratix to assign Job, setup and position
            List<ArborExp> lstExp = new List<ArborExp>();

            int aPos = 0; //Arbor position 1 to X
            int aSetupCount = 0;

            // Suffix on Job is current setup on Job
            // First setup Pos=1, but 1st object in List<> = 0
            int aSetup = lstStartPos[aSetupCount];

            try
            {
                foreach (ArborStratix a in lstArborStratix)
                {
                    //Look at each cut in the setup and expand into setup with only single cuts
                    for (int i = 1; i <= a.Nbr; i++)
                    {
                        // Start position counter with 1, then ++ for each pass
                        aPos++;

                        // If there is only on setup start Pos
                        if (lstStartPos.Count == 1)
                            aSetup = 1;
                        else
                        {
                            // Look ahead to start Pos of next setup
                            if (aSetupCount <= lstStartPos.Count)
                            {
                                // If current position = start Pos of next setup, increment the setup count
                                if (lstStartPos[aSetupCount + 1] == aPos)
                                    aSetup++;
                            }
                        }

                        ArborExp aExp = new ArborExp();

                        aExp.Job = string.Concat(job.ToString(), "-", aSetup.ToString());
                        aExp.Wdth = a.Wdth;
                        aExp.Pos = aPos;

                        lstExp.Add(aExp);
                    }
                }
            }
            catch (Exception ex)
            {
                lblError.Text = ex.Message;
                return;
            }

            pBar.Value = 72;

            //Console.WriteLine("==== ARBOR EXPANDED ====");
            //foreach (ArborExp f in lstExp)
            //{
            //    Console.WriteLine(f.Job + " / " + f.Wdth.ToString() + " / " + f.Pos.ToString());
            //}

            int indexClp = 0;

            // Collapse Arbor counting number of repeat sizes as you go
            List<ArborStratix> lstArborClp = new List<ArborStratix>();

            for (int i = 0; i < lstExp.Count(); i++)
            {
                // 1st Position always gets added to the List
                if (lstExp[i].Pos == 1)
                {
                    ArborStratix a = new ArborStratix();

                    a.Job = lstExp[i].Job;
                    a.Wdth = lstExp[i].Wdth;
                    a.Nbr = 1; // Default to get the count started

                    lstArborClp.Add(a);
                }
                else
                {
                    // If Wdth of next element is same as previous
                    if (lstExp[i].Wdth == lstArborClp[indexClp].Wdth)
                    {
                        // lstExp.Wdth is same as previous lstArborClp, so just ++ Nbr
                        lstArborClp[indexClp].Nbr++;
                    }
                    else
                    {
                        // New wdth, so add to lstArborClp
                        ArborStratix a = new ArborStratix();

                        a.Job = lstExp[i].Job;
                        a.Wdth = lstExp[i].Wdth;
                        a.Nbr = 1; // Default to get the count started

                        lstArborClp.Add(a);

                        // Inc index of lstArborClp
                        indexClp++;
                    }
                }
            }

            //Console.WriteLine("==== ARBOR COLLAPSED ====");
            //foreach (ArborStratix f in lstArborClp)
            //{
            //    Console.WriteLine(f.Job + " / " + f.Wdth.ToString() + " / " + f.Nbr.ToString());
            //}

            // Build MultFile
            List<MultFile> lstMults = new List<MultFile>();

            foreach (ArborStratix a in lstArborClp)
            {
                MultFile m = new MultFile();

                m.Job = a.Job;
                m.Qty = a.Nbr;
                m.Size = a.Wdth;
                m.WdthP = Convert.ToDecimal(lstDetail.Where(x => x.Wdth == a.Wdth).Select(x => x.WdthP).FirstOrDefault());
                m.WdthN = Convert.ToDecimal(lstDetail.Where(x => x.Wdth == a.Wdth).Select(x => x.WdthN).FirstOrDefault());

                lstMults.Add(m);
            }

            pBar.Value = 90;

            // testing
            Console.WriteLine("==== MULT FILE ====");
            foreach (MultFile h in lstMults)
            {
                Console.WriteLine(h.Job + " / " + h.Cust + " / " + h.Qty.ToString() + " / " + h.Size.ToString() + " / " + h.WdthP.ToString() + " / " + h.WdthN.ToString() + " / " + h.Knife);
            }
            #endregion

            #region Exports
            ExcelExport objXLS = new ExcelExport();

            try
            {
                objXLS.WriteHdr(lstHdr);
            }
            catch (Exception ex)
            {
                lblError.Text = ex.Message;
                return;
            }

            try
            {
                objXLS.WriteMults(lstMults);
            }
            catch (Exception ex)
            {
                lblError.Text = ex.Message;
                return;
            }
            #endregion

            // Write Hdr and Mults to ListViews
            try
            {
                ListView_Fill(lstHdr, lstMults);
            }
            catch (Exception ex)
            {
                lblError.Text = ex.Message;
                return;
            }

            // Progress complete
            pBar.Value = 100;

            // Show where files were written
            string hdrFileName = ConfigurationManager.AppSettings.Get("HdrFileName");
            string multFileName = ConfigurationManager.AppSettings.Get("MultFileName");
            string destPath = ConfigurationManager.AppSettings.Get("DestPath");

            lblFiles.Text = "Files written:\n" + Path.Combine(destPath, hdrFileName) + "\n" + Path.Combine(destPath, multFileName);

        }

        private void ListView_Fill(List<HdrFile> lstHdr, List<MultFile> lstMults)
        {
            // lvHeader
            ColumnHeader hJob, hMtl, hWdth, hGa, hPerc, hClr, hGaP, hGaN, hNote;

            hJob = new ColumnHeader();
            hJob.Text = "Job-Setup";
            hJob.TextAlign = HorizontalAlignment.Left;
            hJob.Width = 75;
        
            hMtl = new ColumnHeader();
            hMtl.Text = "Mtl";
            hMtl.TextAlign = HorizontalAlignment.Left;
            hMtl.Width = 40;

            hWdth = new ColumnHeader();
            hWdth.Text = "Wdth";
            hWdth.TextAlign = HorizontalAlignment.Left;
            hWdth.Width = 60;

            hGa = new ColumnHeader();
            hGa.Text = "Ga";
            hGa.TextAlign = HorizontalAlignment.Left;
            hGa.Width = 60;

            hPerc = new ColumnHeader();
            hPerc.Text = "Perc";
            hPerc.TextAlign = HorizontalAlignment.Left;
            hPerc.Width = 40;

            hClr = new ColumnHeader();
            hClr.Text = "Clr";
            hClr.TextAlign = HorizontalAlignment.Left;
            hClr.Width = 60;

            hGaP = new ColumnHeader();
            hGaP.Text = "Ga(+)";
            hGaP.TextAlign = HorizontalAlignment.Left;
            hGaP.Width = 60;

            hGaN = new ColumnHeader();
            hGaN.Text = "Ga(-)";
            hGaN.TextAlign = HorizontalAlignment.Left;
            hGaN.Width = 60;

            hNote = new ColumnHeader();
            hNote.Text = "Note";
            hNote.TextAlign = HorizontalAlignment.Left;
            hNote.Width = 100;

            lvHeader.View = View.Details;

            lvHeader.Columns.Add(hJob);
            lvHeader.Columns.Add(hMtl);
            lvHeader.Columns.Add(hWdth);
            lvHeader.Columns.Add(hGa);
            lvHeader.Columns.Add(hPerc);
            lvHeader.Columns.Add(hClr);
            lvHeader.Columns.Add(hGaP);
            lvHeader.Columns.Add(hGaN);
            lvHeader.Columns.Add(hNote);

            lvHeader.Items.Clear();

            foreach (HdrFile h in lstHdr)
            {
                ListViewItem lvi = new ListViewItem();

                // Need to put the 1st column in Text or row starts in 2nd column
                lvi.Text = h.Job;
                lvi.SubItems.Add(h.Mtl);
                lvi.SubItems.Add(h.Wdth.ToString());
                lvi.SubItems.Add(h.Ga.ToString().TrimEnd('0'));
                lvi.SubItems.Add(h.KnifeClr.ToString());
                lvi.SubItems.Add(h.Clr.ToString().TrimEnd('0'));
                lvi.SubItems.Add(h.GaP.ToString().TrimEnd('0'));
                lvi.SubItems.Add(h.GaN.ToString().TrimEnd('0'));
                lvi.SubItems.Add(h.Note);

                lvHeader.Items.Add(lvi);
            }

            // lvMults
            ColumnHeader mJob, mQty, mWdth, mWdthP, mWdthN;

            mJob = new ColumnHeader();
            mJob.Text = "Job-Setup";
            mJob.TextAlign = HorizontalAlignment.Left;
            mJob.Width = 75;

            mQty = new ColumnHeader();
            mQty.Text = "Qty";
            mQty.TextAlign = HorizontalAlignment.Left;
            mQty.Width = 40;

            mWdth = new ColumnHeader();
            mWdth.Text = "Wdth";
            mWdth.TextAlign = HorizontalAlignment.Left;
            mWdth.Width = 60;

            mWdthP = new ColumnHeader();
            mWdthP.Text = "WdthP";
            mWdthP.TextAlign = HorizontalAlignment.Left;
            mWdthP.Width = 60;

            mWdthN = new ColumnHeader();
            mWdthN.Text = "WdthN";
            mWdthN.TextAlign = HorizontalAlignment.Left;
            mWdthN.Width = 60;

            lvMults.View = View.Details;
            lvMults.Columns.Add(mJob);
            lvMults.Columns.Add(mQty);
            lvMults.Columns.Add(mWdth);
            lvMults.Columns.Add(mWdthP);
            lvMults.Columns.Add(mWdthN);

            lvMults.Items.Clear();

            foreach (MultFile m in lstMults)
            {
                ListViewItem lvi = new ListViewItem();

                lvi.Text = m.Job;
                lvi.SubItems.Add(m.Qty.ToString());
                lvi.SubItems.Add(m.Size.ToString());
                lvi.SubItems.Add(m.WdthP.ToString().TrimEnd('0'));
                lvi.SubItems.Add(m.WdthN.ToString().TrimEnd('0'));

                lvMults.Items.Add(lvi);
            }
        }

        private void txtJob_Validating(object sender, CancelEventArgs e)
        {
            bool cancel = false;

            if (Helpers.IsInteger(txtJob.Text))
            {
                // Control passed validation
                cancel = false;
            }
            else
            {
                //This control has failed validation
                cancel = true;
                lblError.Text = "Job is not a number";
                //this.errorProvider1.SetError(this.txtWidth, "Enter Width with 3 decimals");
            }

            e.Cancel = cancel;
        }
    }
}
