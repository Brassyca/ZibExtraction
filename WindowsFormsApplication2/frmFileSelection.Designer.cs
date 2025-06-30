namespace Zibs.ZibExtraction
{
    partial class frmFileSelection
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
            this.lbImageLocation = new System.Windows.Forms.Label();
            this.tbImageLocation = new System.Windows.Forms.TextBox();
            this.btImageLocation = new System.Windows.Forms.Button();
            this.lbCommonConfigLocation = new System.Windows.Forms.Label();
            this.lbCodeSystemCodesLocation = new System.Windows.Forms.Label();
            this.tbCodeSystemCodesLocation = new System.Windows.Forms.TextBox();
            this.tbCommonConfigLocation = new System.Windows.Forms.TextBox();
            this.btCodeSystemCodesLocation = new System.Windows.Forms.Button();
            this.btCommonConfigLocation = new System.Windows.Forms.Button();
            this.btCancel = new System.Windows.Forms.Button();
            this.btOK = new System.Windows.Forms.Button();
            this.gbHelp = new System.Windows.Forms.GroupBox();
            this.lbHelp = new System.Windows.Forms.Label();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.gbHelp.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // lbImageLocation
            // 
            this.lbImageLocation.AutoSize = true;
            this.lbImageLocation.Location = new System.Drawing.Point(12, 48);
            this.lbImageLocation.Name = "lbImageLocation";
            this.lbImageLocation.Size = new System.Drawing.Size(80, 13);
            this.lbImageLocation.TabIndex = 0;
            this.lbImageLocation.Text = "ImageLocation:";
            // 
            // tbImageLocation
            // 
            this.tbImageLocation.Location = new System.Drawing.Point(160, 45);
            this.tbImageLocation.Name = "tbImageLocation";
            this.tbImageLocation.Size = new System.Drawing.Size(547, 20);
            this.tbImageLocation.TabIndex = 1;
            // 
            // btImageLocation
            // 
            this.btImageLocation.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btImageLocation.AutoSize = true;
            this.btImageLocation.Location = new System.Drawing.Point(713, 43);
            this.btImageLocation.Name = "btImageLocation";
            this.btImageLocation.Size = new System.Drawing.Size(75, 23);
            this.btImageLocation.TabIndex = 2;
            this.btImageLocation.Text = "Bladeren...";
            this.btImageLocation.UseVisualStyleBackColor = true;
            this.btImageLocation.Click += new System.EventHandler(this.btImageLocation_Click);
            // 
            // lbCommonConfigLocation
            // 
            this.lbCommonConfigLocation.AutoSize = true;
            this.lbCommonConfigLocation.Location = new System.Drawing.Point(12, 20);
            this.lbCommonConfigLocation.Name = "lbCommonConfigLocation";
            this.lbCommonConfigLocation.Size = new System.Drawing.Size(122, 13);
            this.lbCommonConfigLocation.TabIndex = 3;
            this.lbCommonConfigLocation.Text = "CommonConfigLocation:";
            // 
            // lbCodeSystemCodesLocation
            // 
            this.lbCodeSystemCodesLocation.AutoSize = true;
            this.lbCodeSystemCodesLocation.Location = new System.Drawing.Point(12, 76);
            this.lbCodeSystemCodesLocation.Name = "lbCodeSystemCodesLocation";
            this.lbCodeSystemCodesLocation.Size = new System.Drawing.Size(140, 13);
            this.lbCodeSystemCodesLocation.TabIndex = 4;
            this.lbCodeSystemCodesLocation.Text = "CodeSystemCodesLocation:";
            // 
            // tbCodeSystemCodesLocation
            // 
            this.tbCodeSystemCodesLocation.Location = new System.Drawing.Point(160, 73);
            this.tbCodeSystemCodesLocation.Name = "tbCodeSystemCodesLocation";
            this.tbCodeSystemCodesLocation.Size = new System.Drawing.Size(547, 20);
            this.tbCodeSystemCodesLocation.TabIndex = 5;
            // 
            // tbCommonConfigLocation
            // 
            this.tbCommonConfigLocation.Location = new System.Drawing.Point(160, 17);
            this.tbCommonConfigLocation.Name = "tbCommonConfigLocation";
            this.tbCommonConfigLocation.Size = new System.Drawing.Size(547, 20);
            this.tbCommonConfigLocation.TabIndex = 6;
            // 
            // btCodeSystemCodesLocation
            // 
            this.btCodeSystemCodesLocation.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btCodeSystemCodesLocation.AutoSize = true;
            this.btCodeSystemCodesLocation.Location = new System.Drawing.Point(713, 71);
            this.btCodeSystemCodesLocation.Name = "btCodeSystemCodesLocation";
            this.btCodeSystemCodesLocation.Size = new System.Drawing.Size(75, 23);
            this.btCodeSystemCodesLocation.TabIndex = 7;
            this.btCodeSystemCodesLocation.Text = "Bladeren...";
            this.btCodeSystemCodesLocation.UseVisualStyleBackColor = true;
            this.btCodeSystemCodesLocation.Click += new System.EventHandler(this.btCodeSystemCodesLocation_Click);
            // 
            // btCommonConfigLocation
            // 
            this.btCommonConfigLocation.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btCommonConfigLocation.AutoSize = true;
            this.btCommonConfigLocation.Location = new System.Drawing.Point(713, 15);
            this.btCommonConfigLocation.Name = "btCommonConfigLocation";
            this.btCommonConfigLocation.Size = new System.Drawing.Size(75, 23);
            this.btCommonConfigLocation.TabIndex = 8;
            this.btCommonConfigLocation.Text = "Bladeren...";
            this.btCommonConfigLocation.UseVisualStyleBackColor = true;
            this.btCommonConfigLocation.Click += new System.EventHandler(this.btCommonConfigLocation_Click);
            // 
            // btCancel
            // 
            this.btCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btCancel.AutoSize = true;
            this.btCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btCancel.Location = new System.Drawing.Point(713, 173);
            this.btCancel.Name = "btCancel";
            this.btCancel.Size = new System.Drawing.Size(75, 23);
            this.btCancel.TabIndex = 9;
            this.btCancel.Text = "Annuleren";
            this.btCancel.UseVisualStyleBackColor = true;
            this.btCancel.Click += new System.EventHandler(this.btCancel_Click);
            // 
            // btOK
            // 
            this.btOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btOK.AutoSize = true;
            this.btOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btOK.Location = new System.Drawing.Point(632, 173);
            this.btOK.Name = "btOK";
            this.btOK.Size = new System.Drawing.Size(75, 23);
            this.btOK.TabIndex = 10;
            this.btOK.Text = "Opslaan";
            this.btOK.UseVisualStyleBackColor = true;
            this.btOK.Click += new System.EventHandler(this.btOK_Click);
            // 
            // gbHelp
            // 
            this.gbHelp.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbHelp.AutoSize = true;
            this.gbHelp.Controls.Add(this.lbHelp);
            this.gbHelp.Location = new System.Drawing.Point(27, 100);
            this.gbHelp.Name = "gbHelp";
            this.gbHelp.Size = new System.Drawing.Size(588, 96);
            this.gbHelp.TabIndex = 11;
            this.gbHelp.TabStop = false;
            this.gbHelp.Text = "Uitleg";
            // 
            // lbHelp
            // 
            this.lbHelp.AutoSize = true;
            this.lbHelp.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbHelp.Location = new System.Drawing.Point(3, 16);
            this.lbHelp.Name = "lbHelp";
            this.lbHelp.Size = new System.Drawing.Size(0, 13);
            this.lbHelp.TabIndex = 0;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel});
            this.statusStrip1.Location = new System.Drawing.Point(0, 199);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(800, 22);
            this.statusStrip1.TabIndex = 12;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel
            // 
            this.toolStripStatusLabel.Name = "toolStripStatusLabel";
            this.toolStripStatusLabel.Size = new System.Drawing.Size(0, 17);
            // 
            // frmFileSelection
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(800, 221);
            this.ControlBox = false;
            this.Controls.Add(this.gbHelp);
            this.Controls.Add(this.btOK);
            this.Controls.Add(this.btCancel);
            this.Controls.Add(this.btCommonConfigLocation);
            this.Controls.Add(this.btCodeSystemCodesLocation);
            this.Controls.Add(this.tbCommonConfigLocation);
            this.Controls.Add(this.tbCodeSystemCodesLocation);
            this.Controls.Add(this.lbCodeSystemCodesLocation);
            this.Controls.Add(this.lbCommonConfigLocation);
            this.Controls.Add(this.btImageLocation);
            this.Controls.Add(this.tbImageLocation);
            this.Controls.Add(this.lbImageLocation);
            this.Controls.Add(this.statusStrip1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "frmFileSelection";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Selectie lege of niet bestaande folder locaties";
            this.Load += new System.EventHandler(this.frmFileSelection_Load);
            this.gbHelp.ResumeLayout(false);
            this.gbHelp.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbImageLocation;
        private System.Windows.Forms.TextBox tbImageLocation;
        private System.Windows.Forms.Button btImageLocation;
        private System.Windows.Forms.Label lbCommonConfigLocation;
        private System.Windows.Forms.Label lbCodeSystemCodesLocation;
        private System.Windows.Forms.TextBox tbCodeSystemCodesLocation;
        private System.Windows.Forms.TextBox tbCommonConfigLocation;
        private System.Windows.Forms.Button btCodeSystemCodesLocation;
        private System.Windows.Forms.Button btCommonConfigLocation;
        private System.Windows.Forms.Button btCancel;
        private System.Windows.Forms.Button btOK;
        private System.Windows.Forms.GroupBox gbHelp;
        private System.Windows.Forms.Label lbHelp;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel;
    }
}