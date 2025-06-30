using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Zibs.Configuration;
using Zibs.ExtensionClasses;

namespace Zibs.ZibExtraction
{
    public partial class frmFileSelection : Form
    {
        private bool forceInput;

        public frmFileSelection(bool _forceInput)
        {
            InitializeComponent();
            forceInput = _forceInput;
            if (forceInput) this.Text = "Herstel folder locaties op verzoek van gebruiker";

            lbHelp.Text= "Bovenstaande locaties dienen gevuld te zijn. Als niet alle drie gevuld zijn zal het programma afsluiten.\r\n" +
                        "Als niet naar de folders verwezen wordt, waar de betreffende bestanden staan, zal het programma op enig moment crashen.\r\n" +
                        "Uitleg: \r\n" +
                        "CommonConfigLocation: Locatie publicatieconfiguratie " + Settings.zibRegistryFileName + "\r\n" +
                        "ImageLocation: Locatie plaatjes t.b.v de spreadsheets       \r\n" +
                        "CodeSystemCodesLocation: Locatie " + Settings.application.CodeSystemCodesFilename + " en " + Settings.application.ZibIdentifiersFilename;
        }

        private void frmFileSelection_Load(object sender, EventArgs e)
        {
            toolStripStatusLabel.Text = "Bestand: " + Path.Combine(Settings.pathAppData, Settings.startConfigFileName);
            statusStrip1.Refresh();

            
            tbImageLocation.Text = Settings.userPreferences.ImageLocation + (Directory.Exists(Settings.userPreferences.ImageLocation)? "" : " (Niet bestaande folder)");
            if (!string.IsNullOrWhiteSpace(tbImageLocation.Text) && Directory.Exists(Settings.userPreferences.ImageLocation))
            {
                tbImageLocation.Enabled = false || forceInput;
                btImageLocation.Enabled = false || forceInput;
                btImageLocation.Tag = true;
            }
            else
            {
                btOK.Enabled = false;
                btImageLocation.Tag = false;
            }
            tbCodeSystemCodesLocation.Text = Settings.application.CodeSystemCodesLocation + (Directory.Exists(Settings.application.CodeSystemCodesLocation) ? "" : " (Niet bestaande folder)");
            if (!string.IsNullOrWhiteSpace(tbCodeSystemCodesLocation.Text) && Directory.Exists(Settings.application.CodeSystemCodesLocation))
            {
                tbCodeSystemCodesLocation.Enabled = false || forceInput;
                btCodeSystemCodesLocation.Enabled = false || forceInput;
                tbCodeSystemCodesLocation.Tag = true;
            }
            else
            {
                btOK.Enabled = false;
                tbCodeSystemCodesLocation.Tag = false;
            }
            tbCommonConfigLocation.Text = Settings.application.CommonConfigLocation + (Directory.Exists(Settings.application.CommonConfigLocation) ? "" : " (Niet bestaande folder)");
            if (!string.IsNullOrWhiteSpace(tbCommonConfigLocation.Text) && Directory.Exists(Settings.application.CommonConfigLocation))
            {
                tbCommonConfigLocation.Enabled = false || forceInput;
                btCommonConfigLocation.Enabled = false || forceInput;
                tbCommonConfigLocation.Tag = true;
            }
            else
            {
                btOK.Enabled = false;
                tbCommonConfigLocation.Tag = false;
            }
        }

        private void btCodeSystemCodesLocation_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog f = new FolderBrowserDialog();
            f.ShowNewFolderButton = true;
            f.Description = "Selecteer de folder waar de codesysteem en zib id bestanden staan";
            f.RootFolder = Environment.SpecialFolder.Desktop;
            f.SelectedPath = Settings.application.CommonConfigLocation;
            DialogResult result = f.ShowDialog();
            if (result == DialogResult.OK)
            {
                Settings.application.CodeSystemCodesLocation = f.SelectedPath;
                this.tbCodeSystemCodesLocation.Text = f.SelectedPath.EllipsisFilename(tbCodeSystemCodesLocation.Size, tbCodeSystemCodesLocation.Font);
                tbCodeSystemCodesLocation.Tag = true;
                EnableOkButton();
            }
        }

        private void btImageLocation_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog f = new FolderBrowserDialog();
            f.ShowNewFolderButton = true;
            f.Description = "Selecteer de folder waar de XLSX plaatjes staan";
            f.RootFolder = Environment.SpecialFolder.Desktop;
            f.SelectedPath = Settings.application.CommonConfigLocation;
            DialogResult result = f.ShowDialog();
            if (result == DialogResult.OK)
            {
                Settings.userPreferences.ImageLocation = f.SelectedPath;
                this.tbImageLocation.Text = f.SelectedPath.EllipsisFilename(tbImageLocation.Size, tbImageLocation.Font);
                tbImageLocation.Tag = true;
                EnableOkButton();
            }
        }

        private void btCommonConfigLocation_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog f = new FolderBrowserDialog();
            f.ShowNewFolderButton = true;
            f.Description = "Selecteer de folder waar de publicatie configuratie bestanden staan";
            f.RootFolder = Environment.SpecialFolder.Desktop;
            f.SelectedPath = Settings.application.CommonConfigLocation;
            DialogResult result = f.ShowDialog();
            if (result == DialogResult.OK)
            {
                Settings.application.CommonConfigLocation = f.SelectedPath;
                this.tbCommonConfigLocation.Text = f.SelectedPath.EllipsisFilename(tbCommonConfigLocation.Size, tbCommonConfigLocation.Font);
                tbCommonConfigLocation.Tag = true;
                EnableOkButton();
            }
        }

        private void EnableOkButton()
        {
            /*(28-12-22) if (!(string.IsNullOrWhiteSpace(tbCodeSystemCodesLocation.Text) ||
                            string.IsNullOrWhiteSpace(tbImageLocation.Text) ||
                            string.IsNullOrWhiteSpace(tbCommonConfigLocation.Text)))   */
            if (((bool)tbCodeSystemCodesLocation.Tag) && ((bool)tbImageLocation.Tag) && ((bool)tbCommonConfigLocation.Tag))
            btOK.Enabled = true;
        }

        private void btOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }



}
