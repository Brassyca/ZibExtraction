namespace Zibs.ZibExtraction
{
    partial class frmManageConfigurations
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmManageConfigurations));
            this.btCancel = new System.Windows.Forms.Button();
            this.btDelete = new System.Windows.Forms.Button();
            this.cbBackup = new System.Windows.Forms.CheckBox();
            this.gbConfigs = new System.Windows.Forms.GroupBox();
            this.SuspendLayout();
            // 
            // btCancel
            // 
            resources.ApplyResources(this.btCancel, "btCancel");
            this.btCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btCancel.Name = "btCancel";
            this.btCancel.UseVisualStyleBackColor = true;
            this.btCancel.Click += new System.EventHandler(this.btCancel_Click);
            // 
            // btDelete
            // 
            resources.ApplyResources(this.btDelete, "btDelete");
            this.btDelete.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btDelete.Name = "btDelete";
            this.btDelete.UseVisualStyleBackColor = true;
            this.btDelete.Click += new System.EventHandler(this.btDelete_Click);
            // 
            // cbBackup
            // 
            resources.ApplyResources(this.cbBackup, "cbBackup");
            this.cbBackup.Name = "cbBackup";
            this.cbBackup.UseVisualStyleBackColor = true;
            // 
            // gbConfigs
            // 
            resources.ApplyResources(this.gbConfigs, "gbConfigs");
            this.gbConfigs.BackColor = System.Drawing.SystemColors.Control;
            this.gbConfigs.Name = "gbConfigs";
            this.gbConfigs.TabStop = false;
            // 
            // frmManageConfigurations
            // 
            this.AcceptButton = this.btDelete;
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit;
            resources.ApplyResources(this, "$this");
            this.CancelButton = this.btCancel;
            this.Controls.Add(this.gbConfigs);
            this.Controls.Add(this.cbBackup);
            this.Controls.Add(this.btDelete);
            this.Controls.Add(this.btCancel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmManageConfigurations";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btCancel;
        private System.Windows.Forms.Button btDelete;
        private System.Windows.Forms.CheckBox cbBackup;
        private System.Windows.Forms.GroupBox gbConfigs;
    }
}