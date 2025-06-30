namespace Zibs.ZibExtraction
{
    partial class frmCopyChoice
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
            this.lbMessage = new System.Windows.Forms.Label();
            this.cbAll = new System.Windows.Forms.CheckBox();
            this.btCopy = new System.Windows.Forms.Button();
            this.btSkip = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lbMessage
            // 
            this.lbMessage.BackColor = System.Drawing.SystemColors.Control;
            this.lbMessage.Dock = System.Windows.Forms.DockStyle.Top;
            this.lbMessage.Location = new System.Drawing.Point(0, 0);
            this.lbMessage.Name = "lbMessage";
            this.lbMessage.Padding = new System.Windows.Forms.Padding(9, 4, 9, 0);
            this.lbMessage.Size = new System.Drawing.Size(267, 47);
            this.lbMessage.TabIndex = 0;
            this.lbMessage.Text = "De doellocatie bevat al een bestand met de naam: ";
            // 
            // cbAll
            // 
            this.cbAll.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cbAll.AutoSize = true;
            this.cbAll.Location = new System.Drawing.Point(6, 79);
            this.cbAll.Name = "cbAll";
            this.cbAll.Size = new System.Drawing.Size(263, 17);
            this.cbAll.TabIndex = 1;
            this.cbAll.Text = "Doe dit voor alle dubbele bestanden in deze aktie ";
            this.cbAll.UseVisualStyleBackColor = true;
            // 
            // btCopy
            // 
            this.btCopy.AutoSize = true;
            this.btCopy.Location = new System.Drawing.Point(6, 50);
            this.btCopy.Name = "btCopy";
            this.btCopy.Size = new System.Drawing.Size(75, 23);
            this.btCopy.TabIndex = 2;
            this.btCopy.Text = "Vervangen";
            this.btCopy.UseVisualStyleBackColor = true;
            this.btCopy.Click += new System.EventHandler(this.btCopy_Click);
            // 
            // btSkip
            // 
            this.btSkip.AutoSize = true;
            this.btSkip.Location = new System.Drawing.Point(183, 50);
            this.btSkip.Name = "btSkip";
            this.btSkip.Size = new System.Drawing.Size(75, 23);
            this.btSkip.TabIndex = 3;
            this.btSkip.Text = "Overslaan";
            this.btSkip.UseVisualStyleBackColor = true;
            this.btSkip.Click += new System.EventHandler(this.btSkip_Click);
            // 
            // frmCopyChoice
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(267, 99);
            this.Controls.Add(this.btSkip);
            this.Controls.Add(this.btCopy);
            this.Controls.Add(this.cbAll);
            this.Controls.Add(this.lbMessage);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmCopyChoice";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Bestand kopieren: vervangen of overslaan";
            this.Load += new System.EventHandler(this.frmCopyChoice_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbMessage;
        private System.Windows.Forms.CheckBox cbAll;
        private System.Windows.Forms.Button btCopy;
        private System.Windows.Forms.Button btSkip;
    }
}