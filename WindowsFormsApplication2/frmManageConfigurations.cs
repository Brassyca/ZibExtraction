using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Zibs.Configuration;

namespace Zibs.ZibExtraction
{
    public partial class frmManageConfigurations : Form
    {
        string selectedConfiguration;
        public frmManageConfigurations(string _selectedConfiguration)
        {
            InitializeComponent();
            selectedConfiguration = _selectedConfiguration;
            InitializeCheckboxes();
            ScaleForm();
        }

        public void InitializeCheckboxes()
        {
            int yStep = Scale(20);
            int xPos = Scale(4);
            int yPos =Scale(20);
            int i = -1;
            CheckBox _checkBox;

            Settings.getConfigurations();
            this.Height = Scale(150) + Scale(20) * (Settings.configList.Count-1);


            foreach (string config in Settings.configList.Where(x => x != "default"))
            {
                _checkBox = new CheckBox();
                i++;

                _checkBox.AutoSize = true;
                _checkBox.Location = new Point(xPos, yPos + i* yStep);
                _checkBox.Name = "cb"+ config;
                _checkBox.Text = config;
                _checkBox.UseVisualStyleBackColor = true;
                _checkBox.Enabled = config == selectedConfiguration ? false : true;
                gbConfigs.Controls.Add(_checkBox);
            }
        }

        private void btCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btDelete_Click(object sender, EventArgs e)
        {
            int checkedboxes = 0;
            List<string> configurationsToRemove;
            foreach (CheckBox cb in gbConfigs.Controls.OfType<CheckBox>())
            {
                if (cb.Checked) checkedboxes++;
            }

            if (checkedboxes > 0 )
                {
                DialogResult answ = MessageBox.Show("Je staat op punt " + checkedboxes.ToString() + "configuratie" + (checkedboxes > 1 ? "s" : "")
                    + "\r\nte verwijderen. Doorgaan?", "Configuraties verwijderen", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (answ == DialogResult.OK)
                {
                    configurationsToRemove = new List<string>();
                    foreach (CheckBox cb in gbConfigs.Controls.OfType<CheckBox>().Where(x => x.Checked))
                    {
                        configurationsToRemove.Add(cb.Text);
                    }
                    bool removeSucces = Settings.removeConfigurations(configurationsToRemove, cbBackup.Checked);
                    if (!removeSucces) MessageBox.Show("Het verwijderen van de configuraties is niet gelukt");
                    Settings.getConfigurations();
                }
            }
      
            this.Close();
        }


        private void ScaleForm()
        {
            // DPI maakt de buttons veel te groot
            if (this.AutoScaleMode == AutoScaleMode.Dpi)
            {
                    this.SuspendLayout();
                    // foreach button mag niet met ref
                    Button[] buttons = this.Controls.OfType<Button>().ToArray();
                    for (int i = 0; i < buttons.Count(); i++)
                        ScaleButton(ref buttons[i]);
            }
            gbConfigs.SuspendLayout();
            int maxRbWidth = gbConfigs.Controls.OfType<CheckBox>().Select(x => x.Width).Max();
            int buttonWidth = btDelete.Width + Scale(6) + btCancel.Width + Scale(20);
            if (Math.Max(maxRbWidth, buttonWidth) > gbConfigs.Width)
            {
                int increase = Math.Max(maxRbWidth, buttonWidth) - gbConfigs.Width + Scale(10);
                this.Width += increase;
                //gbConfigs.Width += increase;
            }
            btDelete.Location = new Point(gbConfigs.Width - Scale(10) - btDelete.Width, gbConfigs.Bottom + Scale(6));
            btCancel.Location = new Point(btDelete.Left - Scale(6) - btCancel.Width, btDelete.Top);
            cbBackup.Location = new Point(gbConfigs.Left + Scale(4), btDelete.Bottom + Scale(6));

            gbConfigs.ResumeLayout(false);
            gbConfigs.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
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

