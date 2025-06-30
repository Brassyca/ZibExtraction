namespace Zibs
{
    namespace ZibExtraction
    {
        partial class frmConfigChoice
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
            this.cbConfigs = new System.Windows.Forms.ComboBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // cbConfigs
            // 
            this.cbConfigs.FormattingEnabled = true;
            this.cbConfigs.Location = new System.Drawing.Point(12, 15);
            this.cbConfigs.Margin = new System.Windows.Forms.Padding(2);
            this.cbConfigs.Name = "cbConfigs";
            this.cbConfigs.Size = new System.Drawing.Size(178, 21);
            this.cbConfigs.TabIndex = 0;
            // 
            // btnOK
            // 
            this.btnOK.AutoSize = true;
            this.btnOK.Location = new System.Drawing.Point(131, 53);
            this.btnOK.Margin = new System.Windows.Forms.Padding(2);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(52, 23);
            this.btnOK.TabIndex = 2;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // frmConfigChoice
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(196, 83);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.cbConfigs);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(2);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmConfigChoice";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Kies een configuratie";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmConfigChoice_FormClosing);
            this.Load += new System.EventHandler(this.frmConfigChoice_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

            }

            #endregion

            private System.Windows.Forms.ComboBox cbConfigs;
            private System.Windows.Forms.Button btnOK;
        }
    }
}