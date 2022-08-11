namespace MobileReport
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
            this.metroTabControl1 = new MetroFramework.Controls.MetroTabControl();
            this.metroTabPage1 = new MetroFramework.Controls.MetroTabPage();
            this.lblHome = new System.Windows.Forms.Label();
            this.lstStatus = new System.Windows.Forms.ListBox();
            this.btnCreate = new MetroFramework.Controls.MetroTile();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.lstReference = new System.Windows.Forms.ListBox();
            this.btnAREACODE = new MetroFramework.Controls.MetroTile();
            this.btnGLCODE = new MetroFramework.Controls.MetroTile();
            this.metroTabPage2 = new MetroFramework.Controls.MetroTabPage();
            this.lblDescRog = new System.Windows.Forms.Label();
            this.btnCombinRogers = new MetroFramework.Controls.MetroTile();
            this.btnRogers_IOCC = new MetroFramework.Controls.MetroTile();
            this.btnRogers_CCD = new MetroFramework.Controls.MetroTile();
            this.metroTabPage3 = new MetroFramework.Controls.MetroTabPage();
            this.lblverizonGL = new System.Windows.Forms.Label();
            this.lblVerizon = new System.Windows.Forms.Label();
            this.btnCreateVerizonReport = new MetroFramework.Controls.MetroTile();
            this.btnVerizon = new MetroFramework.Controls.MetroTile();
            this.metroTabPage4 = new MetroFramework.Controls.MetroTabPage();
            this.lblBellGLAR = new System.Windows.Forms.Label();
            this.lblBell = new System.Windows.Forms.Label();
            this.btnBellReport = new MetroFramework.Controls.MetroTile();
            this.btnBellFile = new MetroFramework.Controls.MetroTile();
            this.metroTabPage5 = new MetroFramework.Controls.MetroTabPage();
            this.lblContact = new System.Windows.Forms.Label();
            this.metroTabControl1.SuspendLayout();
            this.metroTabPage1.SuspendLayout();
            this.metroTabPage2.SuspendLayout();
            this.metroTabPage3.SuspendLayout();
            this.metroTabPage4.SuspendLayout();
            this.metroTabPage5.SuspendLayout();
            this.SuspendLayout();
            // 
            // metroTabControl1
            // 
            this.metroTabControl1.Controls.Add(this.metroTabPage1);
            this.metroTabControl1.Controls.Add(this.metroTabPage2);
            this.metroTabControl1.Controls.Add(this.metroTabPage3);
            this.metroTabControl1.Controls.Add(this.metroTabPage4);
            this.metroTabControl1.Controls.Add(this.metroTabPage5);
            this.metroTabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.metroTabControl1.Location = new System.Drawing.Point(0, 60);
            this.metroTabControl1.Name = "metroTabControl1";
            this.metroTabControl1.SelectedIndex = 4;
            this.metroTabControl1.Size = new System.Drawing.Size(416, 326);
            this.metroTabControl1.TabIndex = 0;
            this.metroTabControl1.UseSelectable = true;
            // 
            // metroTabPage1
            // 
            this.metroTabPage1.Controls.Add(this.lblHome);
            this.metroTabPage1.Controls.Add(this.lstStatus);
            this.metroTabPage1.Controls.Add(this.btnCreate);
            this.metroTabPage1.Controls.Add(this.label2);
            this.metroTabPage1.Controls.Add(this.label1);
            this.metroTabPage1.Controls.Add(this.lstReference);
            this.metroTabPage1.Controls.Add(this.btnAREACODE);
            this.metroTabPage1.Controls.Add(this.btnGLCODE);
            this.metroTabPage1.HorizontalScrollbarBarColor = true;
            this.metroTabPage1.HorizontalScrollbarHighlightOnWheel = false;
            this.metroTabPage1.HorizontalScrollbarSize = 10;
            this.metroTabPage1.Location = new System.Drawing.Point(4, 38);
            this.metroTabPage1.Name = "metroTabPage1";
            this.metroTabPage1.Size = new System.Drawing.Size(408, 284);
            this.metroTabPage1.TabIndex = 0;
            this.metroTabPage1.Text = "Home";
            this.metroTabPage1.VerticalScrollbarBarColor = true;
            this.metroTabPage1.VerticalScrollbarHighlightOnWheel = false;
            this.metroTabPage1.VerticalScrollbarSize = 3;
            // 
            // lblHome
            // 
            this.lblHome.BackColor = System.Drawing.Color.White;
            this.lblHome.Location = new System.Drawing.Point(236, 87);
            this.lblHome.Name = "lblHome";
            this.lblHome.Size = new System.Drawing.Size(169, 90);
            this.lblHome.TabIndex = 8;
            // 
            // lstStatus
            // 
            this.lstStatus.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.lstStatus.Font = new System.Drawing.Font("Arial Narrow", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstStatus.FormattingEnabled = true;
            this.lstStatus.ItemHeight = 16;
            this.lstStatus.Location = new System.Drawing.Point(160, 194);
            this.lstStatus.Name = "lstStatus";
            this.lstStatus.Size = new System.Drawing.Size(245, 80);
            this.lstStatus.TabIndex = 7;
            // 
            // btnCreate
            // 
            this.btnCreate.ActiveControl = null;
            this.btnCreate.Location = new System.Drawing.Point(6, 190);
            this.btnCreate.Name = "btnCreate";
            this.btnCreate.Size = new System.Drawing.Size(148, 85);
            this.btnCreate.TabIndex = 6;
            this.btnCreate.Text = "Generate \r\nIntegrated Report";
            this.btnCreate.TileTextFontWeight = MetroFramework.MetroTileTextWeight.Bold;
            this.btnCreate.UseSelectable = true;
            this.btnCreate.Click += new System.EventHandler(this.btnCreate_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.White;
            this.label2.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(8, 162);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(110, 15);
            this.label2.TabIndex = 5;
            this.label2.Text = "Create a Report";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.White;
            this.label1.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(8, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(146, 15);
            this.label1.TabIndex = 4;
            this.label1.Text = "Upload Reference file";
            // 
            // lstReference
            // 
            this.lstReference.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.lstReference.FormattingEnabled = true;
            this.lstReference.Location = new System.Drawing.Point(236, 45);
            this.lstReference.Name = "lstReference";
            this.lstReference.Size = new System.Drawing.Size(169, 39);
            this.lstReference.TabIndex = 0;
            // 
            // btnAREACODE
            // 
            this.btnAREACODE.ActiveControl = null;
            this.btnAREACODE.Location = new System.Drawing.Point(121, 45);
            this.btnAREACODE.Name = "btnAREACODE";
            this.btnAREACODE.Size = new System.Drawing.Size(109, 85);
            this.btnAREACODE.TabIndex = 3;
            this.btnAREACODE.Text = "AREA CODE";
            this.btnAREACODE.TileTextFontWeight = MetroFramework.MetroTileTextWeight.Bold;
            this.btnAREACODE.UseSelectable = true;
            this.btnAREACODE.Click += new System.EventHandler(this.btnAREACODE_Click);
            // 
            // btnGLCODE
            // 
            this.btnGLCODE.ActiveControl = null;
            this.btnGLCODE.Location = new System.Drawing.Point(6, 45);
            this.btnGLCODE.Name = "btnGLCODE";
            this.btnGLCODE.Size = new System.Drawing.Size(109, 85);
            this.btnGLCODE.TabIndex = 2;
            this.btnGLCODE.Text = "GL CODE";
            this.btnGLCODE.TileTextFontWeight = MetroFramework.MetroTileTextWeight.Bold;
            this.btnGLCODE.UseSelectable = true;
            this.btnGLCODE.Click += new System.EventHandler(this.btnGLCODE_Click);
            // 
            // metroTabPage2
            // 
            this.metroTabPage2.BackColor = System.Drawing.SystemColors.Control;
            this.metroTabPage2.Controls.Add(this.lblDescRog);
            this.metroTabPage2.Controls.Add(this.btnCombinRogers);
            this.metroTabPage2.Controls.Add(this.btnRogers_IOCC);
            this.metroTabPage2.Controls.Add(this.btnRogers_CCD);
            this.metroTabPage2.HorizontalScrollbarBarColor = true;
            this.metroTabPage2.HorizontalScrollbarHighlightOnWheel = false;
            this.metroTabPage2.HorizontalScrollbarSize = 10;
            this.metroTabPage2.Location = new System.Drawing.Point(4, 38);
            this.metroTabPage2.Name = "metroTabPage2";
            this.metroTabPage2.Size = new System.Drawing.Size(408, 284);
            this.metroTabPage2.TabIndex = 1;
            this.metroTabPage2.Text = "Rogers";
            this.metroTabPage2.VerticalScrollbarBarColor = true;
            this.metroTabPage2.VerticalScrollbarHighlightOnWheel = false;
            this.metroTabPage2.VerticalScrollbarSize = 3;
            // 
            // lblDescRog
            // 
            this.lblDescRog.BackColor = System.Drawing.Color.Navy;
            this.lblDescRog.ForeColor = System.Drawing.SystemColors.Control;
            this.lblDescRog.Location = new System.Drawing.Point(23, 134);
            this.lblDescRog.Name = "lblDescRog";
            this.lblDescRog.Size = new System.Drawing.Size(254, 145);
            this.lblDescRog.TabIndex = 5;
            // 
            // btnCombinRogers
            // 
            this.btnCombinRogers.ActiveControl = null;
            this.btnCombinRogers.Location = new System.Drawing.Point(288, 29);
            this.btnCombinRogers.Name = "btnCombinRogers";
            this.btnCombinRogers.Size = new System.Drawing.Size(108, 252);
            this.btnCombinRogers.Style = MetroFramework.MetroColorStyle.Purple;
            this.btnCombinRogers.TabIndex = 4;
            this.btnCombinRogers.Text = "Rogers \r\nReport";
            this.btnCombinRogers.TileTextFontWeight = MetroFramework.MetroTileTextWeight.Bold;
            this.btnCombinRogers.UseSelectable = true;
            this.btnCombinRogers.Click += new System.EventHandler(this.btnCombinRogers_Click);
            this.btnCombinRogers.MouseLeave += new System.EventHandler(this.btnCombinRogers_MouseLeave);
            this.btnCombinRogers.MouseHover += new System.EventHandler(this.btnCombinRogers_MouseHover);
            // 
            // btnRogers_IOCC
            // 
            this.btnRogers_IOCC.ActiveControl = null;
            this.btnRogers_IOCC.Location = new System.Drawing.Point(153, 29);
            this.btnRogers_IOCC.Name = "btnRogers_IOCC";
            this.btnRogers_IOCC.Size = new System.Drawing.Size(124, 90);
            this.btnRogers_IOCC.Style = MetroFramework.MetroColorStyle.Orange;
            this.btnRogers_IOCC.TabIndex = 3;
            this.btnRogers_IOCC.Text = "IOCC File";
            this.btnRogers_IOCC.TileTextFontWeight = MetroFramework.MetroTileTextWeight.Bold;
            this.btnRogers_IOCC.UseSelectable = true;
            this.btnRogers_IOCC.Click += new System.EventHandler(this.btnRogers_IOCC_Click);
            this.btnRogers_IOCC.MouseLeave += new System.EventHandler(this.btnRogers_IOCC_MouseLeave);
            this.btnRogers_IOCC.MouseHover += new System.EventHandler(this.btnRogers_IOCC_MouseHover);
            // 
            // btnRogers_CCD
            // 
            this.btnRogers_CCD.ActiveControl = null;
            this.btnRogers_CCD.Location = new System.Drawing.Point(23, 29);
            this.btnRogers_CCD.Name = "btnRogers_CCD";
            this.btnRogers_CCD.Size = new System.Drawing.Size(124, 90);
            this.btnRogers_CCD.Style = MetroFramework.MetroColorStyle.Green;
            this.btnRogers_CCD.TabIndex = 2;
            this.btnRogers_CCD.Text = "CCD File";
            this.btnRogers_CCD.TileTextFontWeight = MetroFramework.MetroTileTextWeight.Bold;
            this.btnRogers_CCD.UseSelectable = true;
            this.btnRogers_CCD.Click += new System.EventHandler(this.btnRogers_CCD_Click);
            this.btnRogers_CCD.MouseLeave += new System.EventHandler(this.btnRogers_CCD_MouseLeave);
            this.btnRogers_CCD.MouseHover += new System.EventHandler(this.btnRogers_CCD_MouseHover);
            // 
            // metroTabPage3
            // 
            this.metroTabPage3.Controls.Add(this.lblverizonGL);
            this.metroTabPage3.Controls.Add(this.lblVerizon);
            this.metroTabPage3.Controls.Add(this.btnCreateVerizonReport);
            this.metroTabPage3.Controls.Add(this.btnVerizon);
            this.metroTabPage3.HorizontalScrollbarBarColor = true;
            this.metroTabPage3.HorizontalScrollbarHighlightOnWheel = false;
            this.metroTabPage3.HorizontalScrollbarSize = 10;
            this.metroTabPage3.Location = new System.Drawing.Point(4, 38);
            this.metroTabPage3.Name = "metroTabPage3";
            this.metroTabPage3.Size = new System.Drawing.Size(408, 284);
            this.metroTabPage3.TabIndex = 2;
            this.metroTabPage3.Text = "Verizon";
            this.metroTabPage3.VerticalScrollbarBarColor = true;
            this.metroTabPage3.VerticalScrollbarHighlightOnWheel = false;
            this.metroTabPage3.VerticalScrollbarSize = 3;
            // 
            // lblverizonGL
            // 
            this.lblverizonGL.BackColor = System.Drawing.Color.White;
            this.lblverizonGL.Location = new System.Drawing.Point(24, 184);
            this.lblverizonGL.Name = "lblverizonGL";
            this.lblverizonGL.Size = new System.Drawing.Size(171, 92);
            this.lblverizonGL.TabIndex = 7;
            // 
            // lblVerizon
            // 
            this.lblVerizon.BackColor = System.Drawing.Color.White;
            this.lblVerizon.Location = new System.Drawing.Point(201, 21);
            this.lblVerizon.Name = "lblVerizon";
            this.lblVerizon.Size = new System.Drawing.Size(204, 157);
            this.lblVerizon.TabIndex = 6;
            // 
            // btnCreateVerizonReport
            // 
            this.btnCreateVerizonReport.ActiveControl = null;
            this.btnCreateVerizonReport.Location = new System.Drawing.Point(201, 184);
            this.btnCreateVerizonReport.Name = "btnCreateVerizonReport";
            this.btnCreateVerizonReport.Size = new System.Drawing.Size(109, 92);
            this.btnCreateVerizonReport.Style = MetroFramework.MetroColorStyle.Silver;
            this.btnCreateVerizonReport.TabIndex = 3;
            this.btnCreateVerizonReport.Text = "Verizon \r\nReport";
            this.btnCreateVerizonReport.TileTextFontWeight = MetroFramework.MetroTileTextWeight.Bold;
            this.btnCreateVerizonReport.UseSelectable = true;
            this.btnCreateVerizonReport.Click += new System.EventHandler(this.btnCreateVerizonReport_Click);
            // 
            // btnVerizon
            // 
            this.btnVerizon.ActiveControl = null;
            this.btnVerizon.Location = new System.Drawing.Point(24, 21);
            this.btnVerizon.Name = "btnVerizon";
            this.btnVerizon.Size = new System.Drawing.Size(171, 157);
            this.btnVerizon.Style = MetroFramework.MetroColorStyle.Black;
            this.btnVerizon.TabIndex = 2;
            this.btnVerizon.Text = "Verizon \r\nOverview Charges \r\nReport";
            this.btnVerizon.TileTextFontWeight = MetroFramework.MetroTileTextWeight.Bold;
            this.btnVerizon.UseSelectable = true;
            this.btnVerizon.Click += new System.EventHandler(this.btnVerizon_Click);
            this.btnVerizon.MouseLeave += new System.EventHandler(this.btnVerizon_MouseLeave);
            this.btnVerizon.MouseHover += new System.EventHandler(this.btnVerizon_MouseHover);
            // 
            // metroTabPage4
            // 
            this.metroTabPage4.Controls.Add(this.lblBellGLAR);
            this.metroTabPage4.Controls.Add(this.lblBell);
            this.metroTabPage4.Controls.Add(this.btnBellReport);
            this.metroTabPage4.Controls.Add(this.btnBellFile);
            this.metroTabPage4.HorizontalScrollbarBarColor = true;
            this.metroTabPage4.HorizontalScrollbarHighlightOnWheel = false;
            this.metroTabPage4.HorizontalScrollbarSize = 10;
            this.metroTabPage4.Location = new System.Drawing.Point(4, 38);
            this.metroTabPage4.Name = "metroTabPage4";
            this.metroTabPage4.Size = new System.Drawing.Size(408, 284);
            this.metroTabPage4.TabIndex = 3;
            this.metroTabPage4.Text = "Bell";
            this.metroTabPage4.VerticalScrollbarBarColor = true;
            this.metroTabPage4.VerticalScrollbarHighlightOnWheel = false;
            this.metroTabPage4.VerticalScrollbarSize = 3;
            // 
            // lblBellGLAR
            // 
            this.lblBellGLAR.BackColor = System.Drawing.Color.White;
            this.lblBellGLAR.Location = new System.Drawing.Point(14, 178);
            this.lblBellGLAR.Name = "lblBellGLAR";
            this.lblBellGLAR.Size = new System.Drawing.Size(171, 92);
            this.lblBellGLAR.TabIndex = 10;
            // 
            // lblBell
            // 
            this.lblBell.BackColor = System.Drawing.Color.White;
            this.lblBell.Location = new System.Drawing.Point(191, 15);
            this.lblBell.Name = "lblBell";
            this.lblBell.Size = new System.Drawing.Size(204, 157);
            this.lblBell.TabIndex = 9;
            // 
            // btnBellReport
            // 
            this.btnBellReport.ActiveControl = null;
            this.btnBellReport.Location = new System.Drawing.Point(191, 178);
            this.btnBellReport.Name = "btnBellReport";
            this.btnBellReport.Size = new System.Drawing.Size(109, 92);
            this.btnBellReport.Style = MetroFramework.MetroColorStyle.Red;
            this.btnBellReport.TabIndex = 8;
            this.btnBellReport.Text = "Bell Report";
            this.btnBellReport.TileTextFontWeight = MetroFramework.MetroTileTextWeight.Bold;
            this.btnBellReport.UseSelectable = true;
            this.btnBellReport.Click += new System.EventHandler(this.btnBellReport_Click);
            // 
            // btnBellFile
            // 
            this.btnBellFile.ActiveControl = null;
            this.btnBellFile.Location = new System.Drawing.Point(14, 15);
            this.btnBellFile.Name = "btnBellFile";
            this.btnBellFile.Size = new System.Drawing.Size(171, 157);
            this.btnBellFile.Style = MetroFramework.MetroColorStyle.Yellow;
            this.btnBellFile.TabIndex = 7;
            this.btnBellFile.Text = "Bell \r\nSubscriber \r\nLevel report";
            this.btnBellFile.TileTextFontWeight = MetroFramework.MetroTileTextWeight.Bold;
            this.btnBellFile.UseSelectable = true;
            this.btnBellFile.Click += new System.EventHandler(this.btnBellFile_Click);
            this.btnBellFile.MouseLeave += new System.EventHandler(this.btnBellFile_MouseLeave);
            this.btnBellFile.MouseHover += new System.EventHandler(this.btnBellFile_MouseHover);
            // 
            // metroTabPage5
            // 
            this.metroTabPage5.Controls.Add(this.lblContact);
            this.metroTabPage5.HorizontalScrollbarBarColor = true;
            this.metroTabPage5.HorizontalScrollbarHighlightOnWheel = false;
            this.metroTabPage5.HorizontalScrollbarSize = 10;
            this.metroTabPage5.Location = new System.Drawing.Point(4, 38);
            this.metroTabPage5.Name = "metroTabPage5";
            this.metroTabPage5.Size = new System.Drawing.Size(408, 284);
            this.metroTabPage5.TabIndex = 4;
            this.metroTabPage5.Text = "Contact";
            this.metroTabPage5.VerticalScrollbarBarColor = true;
            this.metroTabPage5.VerticalScrollbarHighlightOnWheel = false;
            this.metroTabPage5.VerticalScrollbarSize = 8;
            // 
            // lblContact
            // 
            this.lblContact.Location = new System.Drawing.Point(11, 19);
            this.lblContact.Name = "lblContact";
            this.lblContact.Size = new System.Drawing.Size(384, 252);
            this.lblContact.TabIndex = 2;
            this.lblContact.Text = "Creeated by Samuel Jongeun Kim\r\n\r\n- Target framework: .NET Framework 4.8\r\n- Used " +
    "skills: ADO.NET\r\n- Platform target: 64bit\r\n- https://github.com/JKsamuel/MobileR" +
    "eport\r\n\r\n";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(416, 386);
            this.Controls.Add(this.metroTabControl1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.Name = "Form1";
            this.Padding = new System.Windows.Forms.Padding(0, 60, 0, 0);
            this.Style = MetroFramework.MetroColorStyle.Red;
            this.Text = "Mobile Report";
            this.metroTabControl1.ResumeLayout(false);
            this.metroTabPage1.ResumeLayout(false);
            this.metroTabPage1.PerformLayout();
            this.metroTabPage2.ResumeLayout(false);
            this.metroTabPage3.ResumeLayout(false);
            this.metroTabPage4.ResumeLayout(false);
            this.metroTabPage5.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private MetroFramework.Controls.MetroTabControl metroTabControl1;
        private MetroFramework.Controls.MetroTabPage metroTabPage1;
        private MetroFramework.Controls.MetroTabPage metroTabPage2;
        private MetroFramework.Controls.MetroTabPage metroTabPage3;
        private MetroFramework.Controls.MetroTabPage metroTabPage4;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox lstReference;
        private MetroFramework.Controls.MetroTile btnAREACODE;
        private MetroFramework.Controls.MetroTile btnGLCODE;
        private System.Windows.Forms.ListBox lstStatus;
        private MetroFramework.Controls.MetroTile btnCreate;
        private System.Windows.Forms.Label label2;
        private MetroFramework.Controls.MetroTile btnCombinRogers;
        private MetroFramework.Controls.MetroTile btnRogers_IOCC;
        private MetroFramework.Controls.MetroTile btnRogers_CCD;
        private System.Windows.Forms.Label lblDescRog;
        private MetroFramework.Controls.MetroTile btnCreateVerizonReport;
        private MetroFramework.Controls.MetroTile btnVerizon;
        private System.Windows.Forms.Label lblVerizon;
        private System.Windows.Forms.Label lblBell;
        private MetroFramework.Controls.MetroTile btnBellReport;
        private MetroFramework.Controls.MetroTile btnBellFile;
        private System.Windows.Forms.Label lblverizonGL;
        private System.Windows.Forms.Label lblBellGLAR;
        private System.Windows.Forms.Label lblHome;
        private MetroFramework.Controls.MetroTabPage metroTabPage5;
        private System.Windows.Forms.Label lblContact;
    }
}

