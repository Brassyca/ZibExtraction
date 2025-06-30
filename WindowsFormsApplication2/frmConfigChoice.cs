using System;
using System.Windows.Forms;
using System.Drawing;
using Zibs.Configuration;

namespace Zibs
{
    namespace ZibExtraction
    {
        public partial class frmConfigChoice : Form
        {
            string config;

            public string Config
            {
                get { return config; }
            }

            public frmConfigChoice()
            {
                InitializeComponent();
                ScaleForm();
            }

            private void frmConfigChoice_Load(object sender, EventArgs e)
            {
                cbConfigs.Text = "default";
                cbConfigs.DataSource = Settings.configList;
                config = "";
            }

            private void btnOK_Click(object sender, EventArgs e)
            {
                config = cbConfigs.Text;
                this.Close();
            }

            private void btnDefault_Click(object sender, EventArgs e)
            {
                config = "default";
                this.Close();
            }


            private void frmConfigChoice_FormClosing(object sender, FormClosingEventArgs e)
            {
                DialogResult answ = DialogResult.No;
                if (config == "")
                {
                    answ = MessageBox.Show("Er is geen configuratie gekozen \r\nWilt u stoppen?",
                        "Configuratiekeuze", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (answ == DialogResult.No)
                        e.Cancel = true; ;
                }

            }

            private void ScaleForm()
            {
                // Hier evt. schaling
            }

            /// <summary>
            /// Forms schalen niet alle controls even goed. 
            /// Hiermee wordt handmatig de schaling berekend op grond van de AutoScaleBaseSize
            /// </summary>
            /// <param name="parameter">Getal dat geschaald wordt</param>
            /// <returns></returns>
            private int Scale(int parameter)
            {
                Size defaultBaseSize = new Size(5, 13);
                return (int)Math.Round((parameter * AutoScaleBaseSize.Height * 1.0d) / (defaultBaseSize.Height * 1.0d), 0, MidpointRounding.AwayFromZero);
            }
            private void ScaleButton(ref Button button)
            {
                string text = button.Text;
                Font font = button.Font;
                Size size = TextRenderer.MeasureText(text, font);
                button.Size = new Size((int)(1.3 * size.Width), (int)(1.77 * size.Height));
            }

        }
    }
}