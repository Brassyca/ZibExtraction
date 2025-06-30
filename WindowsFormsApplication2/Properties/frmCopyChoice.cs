using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Zibs.ZibExtraction
{
    public partial class frmCopyChoice : Form
    {

        public bool ApplyToAll { get; set; }
        public bool Overwrite { get; set; }
        public string FileName { get; set; }
        public frmCopyChoice()
        {
            InitializeComponent();
        }

        private void btCopy_Click(object sender, EventArgs e)
        {
            this.Overwrite = true;
            if (cbAll.Checked)
                this.ApplyToAll = true;
            else
                this.ApplyToAll = false;
            this.Close();
        }

        private void btSkip_Click(object sender, EventArgs e)
        {
            this.Overwrite = false;
            if (cbAll.Checked)
                this.ApplyToAll = true;
            else
                this.ApplyToAll = false;
            this.Close();
        }

        private void frmCopyChoice_Load(object sender, EventArgs e)
        {

            this.lbMessage.Text += " " + this.FileName;
        }
    }
}
