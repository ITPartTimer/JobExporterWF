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
using JobExporterWF.Log;

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
            this.Text = ConfigurationManager.AppSettings["Title"];

            // Init erro and progress bar
            lblError.Text = "";
            lblFiles.Text = "";
            pBar.Value = 0;

            // Get list of knives from app.config and bind to lvKnives
            var lstKnives = new List<string>(ConfigurationManager.AppSettings["Knives"].Split(new char[] { ';' }));
            var bindingListKnives = new BindingList<decimal>();

            foreach(string k in lstKnives)
            {
                bindingListKnives.Add(Convert.ToDecimal(k));
            }
   
            var sourceKnives = new BindingSource(bindingListKnives, null);

            lvKnives.DataSource = sourceKnives;
        }

        // This method will not execute unless txtJob passes validation
        private void btnExport_Click(object sender, EventArgs e)
        {         
            string job = txtJob.Text;
            decimal knifeDefault = decimal.Parse(lvKnives.GetItemText(lvKnives.SelectedItem));

            Logger.LogWrite("MSG", "Start Job: " + job + " - " + DateTime.Now.ToString());

            // Clear
            lblError.Text = "";
            lblFiles.Text = "";
            pBar.Value = 0;

            // Not sure why, but neither will clear the ListViews
            lvHeader.Clear();
            lvMults.Clear();         

            #region FindJob
            /*********************************
            Find Job in schedule iptpsh_rec.  If not found, a null 
            value is returned with a generic exception message. Format
            error message and return on catch so program does not continue.
            **********************************/
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

                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return");

                return;
            }
            #endregion

            #region Consume
            /*********************************
            Get Planned Consumption from Stratix.
            The number of rows is the number of setups on this Job
            Sort by Seq.  Use Pos information to determine where 
            at what arbor position to start each setup
            *********************************/
            DataAccess objConsume = new DataAccess();

            List<Consume> lstConsume = new List<Consume>();

            try
            {
                lstConsume = objConsume.Get_Consumed(job);
            }
            catch (Exception ex)
            {
                lblError.Text = ex.Message;
                pBar.Value = 0;

                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return");

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

            #region Planned
            /**********************************
            Get Planned Production from Stratix.
            This is a list of unique SOs on the Job.
            Join to CPS table to get width and tolernaces
           **********************************/
            DataAccess objPlanned = new DataAccess();

            List<Planned> lstPlanned = new List<Planned>();

            try
            {
                lstPlanned = objConsume.Get_Planned(job);
            }
            catch (Exception ex)
            {
                lblError.Text = ex.Message;
                pBar.Value = 0;

                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return");

                return;
            }

            pBar.Value = 35;
            #endregion

            #region Ga
            /**********************************
            Get the Gauge from Planned consumption
            
            Determine the most constrained gauge range from the CPS on the Job

            KEVIN uses the gauge and +/- tolerance to create inspection form.
            Stratix creates a most constrained gauge range from the list of SOs
            on the Job.  In this case you know the MIN and MAX allowable gauge.
            Use the Ga from planned consumtion to determine the +/- range.  It 
            might not be even or always positive, but the MIN and MAX will be correct.
            **********************************/

            // Consume query was ordered by seq, so first member contains the TagNo for the job
            string tag = lstConsume.Select(x => x.Tag).First().ToString();
            decimal ga = (decimal)lstConsume.Select(x => x.Ga).First();
          
            if (string.IsNullOrEmpty(tag))
            {
                lblError.Text = "No tag found on Job";
                return;
            }

            DataAccess objGa = new DataAccess();

            Ga g = new Ga();

            try
            {
                g = objGa.Get_Ga(ga, lstPlanned);

                // If Ga is not found, it could be tolling or a closed PO
                if (g.Size == 0)
                    throw new Exception("Ga not found for Job");
            }
            catch (Exception ex)
            {
                lblError.Text = ex.Message;
                pBar.Value = 0;

                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return");

                return;
            }

            pBar.Value = 48;
            #endregion

            #region Build HdrFile          
            /**********************************
            Build List<HdrFile>
            **********************************/

            // Determine number of setups
            int numSetups = lstConsume.Count;
            List<string> lstNumSetups = new List<string>();

            //Add sufix of 1,2,3... to end of Job
            if (numSetups == 1)
                lstNumSetups.Add(string.Concat(job.ToString(), "-", "1"));
            else
                for (int i = 0; i < numSetups; i++)
                    lstNumSetups.Add(string.Concat(job.ToString(), "-", (i + 1).ToString()));


            // lstNumSetups will have same count of members as lstConsume
            int setupCnt = 0;

            List<HdrFile> lstHdr = new List<HdrFile>();

            foreach (string j in lstNumSetups)
            {
                HdrFile h = new HdrFile();

                h.Job = j;
                h.Mtl = lstConsume[0].Frm;
                h.Wdth = lstConsume[0].Wdth;
                h.Ga = g.Size;
                h.Clr = h.KnifeClr * h.Ga;
                h.GaP = g.GaP;
                h.GaN = g.GaN;
                // Get Weight from lstConsume at current position
                h.Wgt = lstConsume[setupCnt].Wgt;

                setupCnt++;

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
            //Console.WriteLine("==== HDR FILE ====");
            //foreach (HdrFile h in lstHdr)
            //{
            //    Console.WriteLine(h.Job + " / " + h.Cust + " / " + h.Mtl + " / " + h.Wdth.ToString() + " / " + h.Ga.ToString() + " / " + h.KnifeClr.ToString() + " / " + h.Clr.ToString() + " / " + h.GaP.ToString() + " / " + h.GaN.ToString() + " / " + h.Note);
            //}
            #endregion

            #region Build MultFile
            /**********************************
            Build MultFile
            **********************************/

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
                pBar.Value = 0;

                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return");

                return;
            }

            // test
            Console.WriteLine("==== STRATIX ARBOR ====");
            foreach (ArborStratix a in lstArborStratix)
            {
                Console.WriteLine(a.Job.ToString() + " / " + a.Wdth.ToString() + " / " + a.Nbr.ToString());
            }

            // Get list of start Pos for each setup
            List<int> lstStartPos = new List<int>();

            foreach (Consume c in lstConsume)
                lstStartPos.Add(c.Pos);

            // Expand lstArborStratix to assign Job, setup and position
            List<ArborExp> lstExp = new List<ArborExp>();

            int aPos = 0; // Arbor position 1 to X
            int aSetupCount = 0; // Keep track of the setup you are on
            int posLookAhead = 0; // Look at next setup position

            int lastSetupPos = lstStartPos[lstStartPos.Count-1];
            int currSetupPos = lstStartPos[aSetupCount];

            // Get value at element 0 in lstStartPos
            // This should always be 1
            int aSetupSuffix = lstStartPos[aSetupCount];

            try
            {   
                // Get each row in lstArborStratix - the way Stratix stores the arbor data in IPTFRA_REC
                foreach (ArborStratix a in lstArborStratix)
                {
                    Console.WriteLine("Before loop - aPos: " + aPos + " / aSetupCount: " + aSetupCount + " / aSetupSuffix: " + aSetupSuffix);

                    //Look at each cut in the setup and insert that cut into List<ArborExp> Nbr times
                    for (int i = 1; i <= a.Nbr; i++)
                    {
                        // Start position counter with 1, then ++ for each pass
                        aPos++;

                        // When currSetupPos = lastSetUpPos, you can no longer look ahead
                        // aSetupSuffix last time through loop wil be used for the rest 
                        // of the cuts in the setup
                        if (lstStartPos.Count == 1)
                            aSetupSuffix = 1;
                        else 
                        {
                            // Keep looking ahead at next setup start position, so long as the current
                            // start position is < the last startup position. If equal, looking ahead
                            // will cause an index out of range error
                            if (currSetupPos < lastSetupPos)
                            {
                                posLookAhead = lstStartPos[aSetupCount + 1];

                                // If current arbor position = start postion of next setup, increment the setup count suffix                                
                                if (posLookAhead == aPos)
                                {
                                    currSetupPos = posLookAhead;

                                    aSetupSuffix++;
                                    aSetupCount++;
                                }
                            }
                        }

                        ArborExp aExp = new ArborExp();

                        aExp.Job = string.Concat(job.ToString(), "-", aSetupSuffix.ToString());
                        aExp.Wdth = a.Wdth;
                        aExp.Pos = aPos;

                        lstExp.Add(aExp);
                    }

                    Console.WriteLine("After loop - aPos: " + aPos + " / aSetupCount: " + aSetupCount + " / aSetupSuffix: " + aSetupSuffix);
                }
            }
            catch (Exception ex)
            {
                lblError.Text = ex.Message;
                pBar.Value = 0;

                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return");

                return;
            }

            pBar.Value = 72;

            Console.WriteLine("==== ARBOR EXPANDED ====");
            foreach (ArborExp f in lstExp)
            {
                Console.WriteLine(f.Job + " / " + f.Wdth.ToString() + " / " + f.Pos.ToString());
            }

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

            Console.WriteLine("==== ARBOR COLLAPSED ====");
            foreach (ArborStratix f in lstArborClp)
            {
                Console.WriteLine(f.Job + " / " + f.Wdth.ToString() + " / " + f.Nbr.ToString());
            }

            // Build MultFile
            List<MultFile> lstMults = new List<MultFile>();

            foreach (ArborStratix a in lstArborClp)
            {
                MultFile m = new MultFile();

                m.Knife = knifeDefault;
                m.Job = a.Job;
                m.Qty = a.Nbr;
                m.Size = a.Wdth;
                m.WdthP = Convert.ToDecimal(lstPlanned.Where(x => x.Wdth == a.Wdth).Select(x => x.WdthP).FirstOrDefault());
                m.WdthN = Convert.ToDecimal(lstPlanned.Where(x => x.Wdth == a.Wdth).Select(x => x.WdthN).FirstOrDefault());

                // If WdthN and WidthP = 0, cut is CTL or a reslit.  Use a default tolerance value.
                if (m.WdthN == 0 && m.WdthP == 0)
                {
                    m.WdthN = Convert.ToDecimal(ConfigurationManager.AppSettings.Get("DefaultWdthT"));
                    m.WdthP = Convert.ToDecimal(ConfigurationManager.AppSettings.Get("DefaultWdthT"));
                }
                    
                lstMults.Add(m);
            }

            pBar.Value = 90;

            // testing
            //Console.WriteLine("==== MULT FILE ====");
            //foreach (MultFile h in lstMults)
            //{
            //    Console.WriteLine(h.Job + " / " + h.Cust + " / " + h.Qty.ToString() + " / " + h.Size.ToString() + " / " + h.WdthP.ToString() + " / " + h.WdthN.ToString() + " / " + h.Knife);
            //}
            #endregion

            #region XLSExports
            /*********************************
            Export Header and Mult List<Object> to XLS
            *********************************/
            ExcelExport objXLS = new ExcelExport();

            try
            {
                objXLS.WriteHdr(lstHdr);
            }
            catch (Exception ex)
            {
                lblError.Text = ex.Message;
                pBar.Value = 0;

                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return");

                return;
            }

            try
            {
                objXLS.WriteMults(lstMults);
            }
            catch (Exception ex)
            {
                lblError.Text = ex.Message;
                pBar.Value = 0;

                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return");

                return;
            }

            try
            {
                objXLS.WriteJobHist(lstHdr,lstMults);
            }
            catch (Exception ex)
            {
                lblError.Text = ex.Message;
                pBar.Value = 0;

                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return");

                return;
            }
            #endregion

            // Write Hdr and Mults to ListViews
            #region FillListViews
            try
            {
                ListView_Fill(lstHdr, lstMults);
            }
            catch (Exception ex)
            {
                lblError.Text = ex.Message;
                pBar.Value = 0;

                Logger.LogWrite("EXC", ex);
                Logger.LogWrite("MSG", "Return");

                return;
            }

            // Progress complete
            pBar.Value = 100;

            // Show where files were written
            string hdrFileName = ConfigurationManager.AppSettings.Get("HdrFileName");
            string multFileName = ConfigurationManager.AppSettings.Get("MultFileName");
            string destPath = ConfigurationManager.AppSettings.Get("DestPath");

            lblFiles.Text = "Files written:\n" + Path.Combine(destPath, hdrFileName) + "\n" + Path.Combine(destPath, multFileName);

            Logger.LogWrite("MSG", "End Job: " + job + " - " + DateTime.Now.ToString());
            #endregion
        }

        private void ListView_Fill(List<HdrFile> lstHdr, List<MultFile> lstMults)
        {
            // lvHeader
            ColumnHeader hJob, hMtl, hWdth, hWgt, hGa, hPerc, hClr, hGaP, hGaN, hNote;

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

            hWgt = new ColumnHeader();
            hWgt.Text = "Wgt";
            hWgt.TextAlign = HorizontalAlignment.Left;
            hWgt.Width = 60;

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
            lvHeader.Columns.Add(hWgt);
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
                lvi.SubItems.Add(h.Wgt.ToString());
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
