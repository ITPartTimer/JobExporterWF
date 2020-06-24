namespace JobExporterWF
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.txtJob = new System.Windows.Forms.TextBox();
            this.btnExport = new System.Windows.Forms.Button();
            this.lvHeader = new System.Windows.Forms.ListView();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lvMults = new System.Windows.Forms.ListView();
            this.lblError = new System.Windows.Forms.Label();
            this.pBar = new System.Windows.Forms.ProgressBar();
            this.lblFiles = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.lvKnives = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(72, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Enter Job No:";
            // 
            // txtJob
            // 
            this.txtJob.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtJob.Location = new System.Drawing.Point(12, 36);
            this.txtJob.Name = "txtJob";
            this.txtJob.Size = new System.Drawing.Size(251, 62);
            this.txtJob.TabIndex = 1;
            this.txtJob.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtJob.Validating += new System.ComponentModel.CancelEventHandler(this.txtJob_Validating);
            // 
            // btnExport
            // 
            this.btnExport.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExport.Location = new System.Drawing.Point(380, 36);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(212, 62);
            this.btnExport.TabIndex = 2;
            this.btnExport.Text = "EXPORT";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // lvHeader
            // 
            this.lvHeader.HideSelection = false;
            this.lvHeader.Location = new System.Drawing.Point(12, 188);
            this.lvHeader.Name = "lvHeader";
            this.lvHeader.Size = new System.Drawing.Size(580, 81);
            this.lvHeader.TabIndex = 3;
            this.lvHeader.UseCompatibleStateImageBehavior = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 172);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(45, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Header:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 272);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(35, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Mults:";
            // 
            // lvMults
            // 
            this.lvMults.HideSelection = false;
            this.lvMults.Location = new System.Drawing.Point(12, 288);
            this.lvMults.Name = "lvMults";
            this.lvMults.Size = new System.Drawing.Size(580, 164);
            this.lvMults.TabIndex = 6;
            this.lvMults.UseCompatibleStateImageBehavior = false;
            // 
            // lblError
            // 
            this.lblError.AutoSize = true;
            this.lblError.ForeColor = System.Drawing.Color.Red;
            this.lblError.Location = new System.Drawing.Point(12, 150);
            this.lblError.Name = "lblError";
            this.lblError.Size = new System.Drawing.Size(39, 13);
            this.lblError.TabIndex = 7;
            this.lblError.Text = "lblError";
            // 
            // pBar
            // 
            this.pBar.Location = new System.Drawing.Point(12, 107);
            this.pBar.Name = "pBar";
            this.pBar.Size = new System.Drawing.Size(251, 23);
            this.pBar.TabIndex = 8;
            // 
            // lblFiles
            // 
            this.lblFiles.AutoSize = true;
            this.lblFiles.Location = new System.Drawing.Point(377, 115);
            this.lblFiles.Name = "lblFiles";
            this.lblFiles.Size = new System.Drawing.Size(38, 13);
            this.lblFiles.TabIndex = 9;
            this.lblFiles.Text = "lblFiles";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(280, 20);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(67, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "Select Knife:";
            // 
            // lvKnives
            // 
            this.lvKnives.FormattingEnabled = true;
            this.lvKnives.Location = new System.Drawing.Point(283, 36);
            this.lvKnives.Name = "lvKnives";
            this.lvKnives.Size = new System.Drawing.Size(77, 95);
            this.lvKnives.TabIndex = 11;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(604, 469);
            this.Controls.Add(this.lvKnives);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.lblFiles);
            this.Controls.Add(this.pBar);
            this.Controls.Add(this.lblError);
            this.Controls.Add(this.lvMults);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lvHeader);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.txtJob);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "STRATIX to KEVIN Exporter";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtJob;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.ListView lvHeader;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ListView lvMults;
        private System.Windows.Forms.Label lblError;
        private System.Windows.Forms.ProgressBar pBar;
        private System.Windows.Forms.Label lblFiles;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ListBox lvKnives;
    }
}

