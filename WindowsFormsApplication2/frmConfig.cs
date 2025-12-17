using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Zibs.Configuration;
using Zibs.ExtensionClasses;

namespace Zibs
{
    namespace ZibExtraction
    {
        public partial class frmConfig : Form
        {
            bool formChanged, XMLLocationChanged, WikiLocationChanged, RTFLocationChanged, XLSLocationChanged, ignoreEvent;
            ZibExtraction parent;
            public delegate void LanguageChangedEventHandler(object sender, EventWithStringArgs e);
            public event LanguageChangedEventHandler LanguageChanged;
            public delegate void ConfigChangedEventHandler(object sender, EventWithStringArgs e);
            public event ConfigChangedEventHandler ConfigChanged;
            public delegate void TemplateChangedEventHandler(object sender, EventWithStringArgs e);
            public event TemplateChangedEventHandler TemplateChanged;
            public delegate void ReleaseChangedEventHandler(object sender, EventWithStringArgs e);
            public event ReleaseChangedEventHandler ReleaseChanged;

            public event EventHandler PrefixChanged;
            bool prefixChanged = false;
            bool configChanged = false;
            string configName;
            bool bitsPasswordChanged = false;
            bool wikiPasswordChanged = false;
            int publicationLabelLeft;

            string ExampleFilesText;



            public bool StartupConfigChanged { get; private set; } = false;
            public bool ImageLocationChanged { get; private set; } = false;
            protected virtual void OnLanguageChanged(EventWithStringArgs e)
            {
                if (LanguageChanged != null)
                {
                    LanguageChanged(this, e);
                }
            }
            protected virtual void OnPrefixChanged(EventArgs e)
            {
                if (PrefixChanged != null)
                {
                    PrefixChanged(this, e);
                }
            }
            protected virtual void OnTemplateChanged(EventWithStringArgs e)
            {
                if (TemplateChanged != null)
                {
                    TemplateChanged(this, e);
                }
            }

            protected virtual void OnReleaseChanged(EventWithStringArgs e)
            {
                if (ReleaseChanged != null)
                {
                    ReleaseChanged(this, e);
                }
            }

            protected virtual void OnConfigChanged(EventWithStringArgs e)
            {
                if (ConfigChanged != null)
                {
                    ConfigChanged(this, e);
                }
            }


            public frmConfig(string config, ZibExtraction _parent)
            {
                InitializeComponent();
                parent = _parent;
                configName = config;
                setFormTitle(configName);
                ScaleForm();
            }

            private void setFormTitle(string _configName)
            {
                this.Text = "Configuratie: " + _configName;

            }

            private void frmConfig_Load(object sender, EventArgs e)
            {
                ConfigToolTip.SetToolTip(tbBitsStatus, "Geef de issuestatusen waarop de query wordt gedaan\r\nScheidt de statussen met ;");


                tbIssuePrefix.Text = Settings.bitscontext.issuePrefix ?? "";
                tbBitsBaseurl.Text = Settings.bitscontext.bitsBaserurl ?? "";
                //            tbBitsPassword.Text = Settings.bitscontext.bitsPassword ?? "";
                tbBitsUser.Text = Settings.bitscontext.bitsUser ?? "";
                tbBitsStatus.Text = Settings.bitscontext.bitsStatus ?? "";
                ExampleFilesText = Settings.userPreferences.ExampleLocation ?? "";
                tbExampleFiles.Text = ExampleFilesText.EllipsisFilename(tbExampleFiles.Size, tbExampleFiles.Font);
                tbMainpage.Text = Settings.wikicontext.MainPage ?? "";
                tBSection.Text = Settings.wikicontext.tocSection.ToString();
                tbLegendpage.Text = Settings.wikicontext.LegendPage ?? "";
                publicationLabelLeft = lblReleaseName.Left;
                List<string> _releaseList = parent.releaseList.Select(x => x[0]).ToList();
                cbReleaseName.DataSource = _releaseList;

                //tbReleaseInfo.Text = Settings.zibcontext.ReleaseInfo ?? "";
                cbReleaseName.Text = Settings.zibcontext.publicatie ?? "";
                tbNumber.Text = Settings.zibcontext.PreReleaseNumber.ToString() ?? "";
                if (Settings.zibcontext.PreReleaseNumber == 0)
                {
                    tbNumber.Visible = false;
                    lblNumber.Visible = false;
                    lblReleaseName.Text = "Publicatie:";
                    lblReleaseName.Left = publicationLabelLeft;
                }
                else
                {
                    tbNumber.Visible = true;
                    lblNumber.Visible = true;
                    lblReleaseName.Text = "Pre-publicatie:";
                    lblReleaseName.Left = publicationLabelLeft - 18;

                }

                tbReleaseInfo.Text = parent.releaseList.Count > 0 ? parent.releaseList[cbReleaseName.SelectedIndex][1] : "";
                tbWikiBaseUrl.Text = Settings.wikicontext.wikiBaserurl ?? "";
                tbWikiFiles.Text = Settings.userPreferences.WikiLocation ?? "";
                tbImages.Text = Settings.userPreferences.ImageLocation ?? "";
                tbWikiUser.Text = Settings.wikicontext.wikiUser ?? "";
                //            tbWikiPassword.Text = Settings.wikicontext.wikiPassword ?? "";
                tbXMLFiles.Text = Settings.userPreferences.XMLLocation ?? "";
                tbXLSFiles.Text = Settings.userPreferences.XLSLocation ?? "";
                tbZibPrefix.Text = Settings.zibcontext.zibPrefix ?? "";
                tbCategory.Text = Settings.zibcontext.zibCategory;
                tbRTFFiles.Text = Settings.userPreferences.ZIB_RTFLocation ?? "";
                ignoreEvent = true;
                tbConfigFile.Text = Settings.application.CommonConfigLocation ?? "";
                tbCodeSystemsFile.Text = Settings.application.CodeSystemCodesLocation ?? "";
                ignoreEvent = false;
                cbTemplate.DataSource = parent.templateList;
                cbTemplate.Text = Settings.zibcontext.zibTemplate ?? "";
                if (Settings.zibcontext.pubLanguage == textLanguage.NL)
                    rbNL.Checked = true;
                else
                    rbEN.Checked = true;

                //            cbTemplate.Text = Settings.zibcontext.zibTemplate;
                //cbWikiStatus.Text = Settings.wikicontext.wikiPagesStatus ?? "";
                lblADRepository.Text = Settings.wikicontext.ArtDecorRepository ?? "";

                formChanged = false;
                prefixChanged = false;
                configChanged = false;
                bitsPasswordChanged = false;
                wikiPasswordChanged = false;
//                configStatusLabel.Text = "Bestand: " + Path.Combine(Settings.pathAppData, Settings.configFileName);
                configStatusLabel.Text = "Bestand: " + Path.Combine(Settings.application.UserConfigLocation, Settings.configFileName);
                ConfigStatus.Refresh();
            }

            private void btnOK_Click(object sender, EventArgs e)
            {
                if (cbNewConfig.Checked && tbConfigName.Text == "")
                {
                    MessageBox.Show("Vul eerst een configuratienaam in");
                    return;
                }

                formChanged = true;
                if (formChanged)
                {
                    if (configChanged)
                    {
                        configName = tbConfigName.Text;
                        OnConfigChanged(new EventWithStringArgs(configName));
                    }
                    Settings.bitscontext.issuePrefix = tbIssuePrefix.Text;
                    Settings.bitscontext.bitsBaserurl = tbBitsBaseurl.Text;
                    Settings.bitscontext.bitsUser = tbBitsUser.Text;

                    if (bitsPasswordChanged && tbBitsPassword.Text.Length != 0) Settings.bitscontext.bitsPassword = tbBitsPassword.Text.Encode();
                    Settings.bitscontext.bitsStatus = tbBitsStatus.Text ?? "";
//                    Settings.userPreferences.ExampleLocation = tbExampleFiles.Text;
                    Settings.userPreferences.ExampleLocation = ExampleFilesText;
                    Settings.userPreferences.ImageLocation = tbImages.Text;
                    Settings.wikicontext.LegendPage = tbLegendpage.Text;
                    Settings.wikicontext.MainPage = tbMainpage.Text;
                    Settings.wikicontext.tocSection = int.Parse(tBSection.Text);
                    Settings.zibcontext.ReleaseInfo = tbReleaseInfo.Text;
                    Settings.zibcontext.publicatie = cbReleaseName.Text;
                    // Toegevoegd 15-12-2025
                    Settings.zibcontext.PreReleaseNumber = int.Parse(tbNumber.Text);

                    Settings.wikicontext.ArtDecorRepository = lblADRepository.Text;
                    Settings.wikicontext.ArtDecorProjectOID = parent.releaseList.Count > 0 ? parent.releaseList[cbReleaseName.SelectedIndex][3] : "";


                    Settings.zibcontext.zibCategory = tbCategory.Text;
                    Settings.wikicontext.wikiBaserurl = tbWikiBaseUrl.Text;
                    if (WikiLocationChanged)
                    {
                        Settings.userPreferences.WikiLocation = tbWikiFiles.Text == "" ? parent.getStandardDirectory("Wiki", Settings.userPreferences.sessionBase) : tbWikiFiles.Text;
                        if (!Directory.Exists(Settings.userPreferences.WikiLocation)) Directory.CreateDirectory(Settings.userPreferences.WikiLocation);
                    }
                    Settings.wikicontext.wikiUser = tbWikiUser.Text;
                    if (wikiPasswordChanged && tbWikiPassword.Text.Length != 0) Settings.wikicontext.wikiPassword = tbWikiPassword.Text.Encode();
                    if (XMLLocationChanged)
                    {
                        Settings.userPreferences.XMLLocation = tbXMLFiles.Text == "" ? parent.getStandardDirectory("XML", Settings.userPreferences.sessionBase) : tbXMLFiles.Text;
                        if (!Directory.Exists(Settings.userPreferences.XMLLocation)) Directory.CreateDirectory(Settings.userPreferences.XMLLocation);
                    }
                    if (XLSLocationChanged)
                    {
                        Settings.userPreferences.XLSLocation = tbXLSFiles.Text == "" ? parent.getStandardDirectory("XLS", Settings.userPreferences.sessionBase) : tbXLSFiles.Text;
                        if (!Directory.Exists(Settings.userPreferences.XLSLocation)) Directory.CreateDirectory(Settings.userPreferences.XLSLocation);
                    }
                    Settings.zibcontext.zibPrefix = tbZibPrefix.Text;

                    if (prefixChanged)
                    {
                        this.OnPrefixChanged(EventArgs.Empty);
                    }
                    Settings.zibcontext.zibTemplate = cbTemplate.Text;

                    //                    Settings.wikicontext.wikiPagesStatus = cbWikiStatus.Text;
                    //                    Settings.wikicontext.ArtDecorRepository = lblADRepository.Text;
                    if (RTFLocationChanged)
                    {
                        Settings.userPreferences.ZIB_RTFLocation = tbRTFFiles.Text == "" ? parent.getStandardDirectory("RTF", Settings.userPreferences.sessionBase) : tbRTFFiles.Text;
                        if (!Directory.Exists(Settings.userPreferences.ZIB_RTFLocation)) Directory.CreateDirectory(Settings.userPreferences.ZIB_RTFLocation);
                    }
                    if (rbNL.Checked)
                        Settings.zibcontext.pubLanguage = textLanguage.NL;
                    else
                        Settings.zibcontext.pubLanguage = textLanguage.EN;
                    try
                    {
                        EventWithStringArgs eStr = new EventWithStringArgs(Settings.zibcontext.pubLanguage.ToString());
                        OnLanguageChanged(eStr);
                    }
                    catch (Exception e2)
                    { MessageBox.Show(e2.Message); }
                }
                this.Close();
            }


            private void btnCancel_Click(object sender, EventArgs e)
            {
                frmConfig_Load(sender, e);
                this.Close();
            }

            private void tbBitsBaseurl_TextChanged(object sender, EventArgs e)
            {
                formChanged = true;
            }

            private void btnExampleFiles_Click(object sender, EventArgs e)
            {
                FolderBrowserDialog f = new FolderBrowserDialog();
                f.ShowNewFolderButton = true;
                f.Description = "Selecteer de folder waar de ZIB voorbeeld bestanden staan";
                f.RootFolder = Environment.SpecialFolder.Desktop;
//                f.SelectedPath = tbExampleFiles.Text;
                f.SelectedPath = ExampleFilesText;
                DialogResult result = f.ShowDialog();
                if (result == DialogResult.OK)
                {
                    ExampleFilesText = f.SelectedPath;
                    this.tbExampleFiles.Text = ExampleFilesText.EllipsisFilename(tbExampleFiles.Size, tbExampleFiles.Font);
                }
            }

            private void btnXMLFiles_Click(object sender, EventArgs e)
            {
                FolderBrowserDialog f = new FolderBrowserDialog();
                f.ShowNewFolderButton = true;
                f.Description = "Selecteer de folder waar de aangemaakte XML bestanden opgeslagen moeten worden";
                f.RootFolder = Environment.SpecialFolder.Desktop;
                f.SelectedPath = this.tbXMLFiles.Text;
                DialogResult result = f.ShowDialog();
                if (result == DialogResult.OK)
                    this.tbXMLFiles.Text = f.SelectedPath;

            }

            private void btnXLSFiles_Click(object sender, EventArgs e)
            {
                FolderBrowserDialog f = new FolderBrowserDialog();
                f.ShowNewFolderButton = true;
                f.Description = "Selecteer de folder waar de aangemaakte spreadsheet bestanden opgeslagen moeten worden";
                f.RootFolder = Environment.SpecialFolder.Desktop;
                f.SelectedPath = this.tbXLSFiles.Text;
                DialogResult result = f.ShowDialog();
                if (result == DialogResult.OK)
                    this.tbXLSFiles.Text = f.SelectedPath;
            }

            private void btnWikiFiles_Click(object sender, EventArgs e)
            {
                FolderBrowserDialog f = new FolderBrowserDialog();
                f.ShowNewFolderButton = true;
                f.Description = "Selecteer de folder waar de aangemaakte Wiki bestanden opgeslagen moeten worden";
                f.RootFolder = Environment.SpecialFolder.Desktop;
                f.SelectedPath = this.tbWikiFiles.Text;
                DialogResult result = f.ShowDialog();
                if (result == DialogResult.OK)
                    this.tbWikiFiles.Text = f.SelectedPath;

            }
            private void btnConfigFile_Click(object sender, EventArgs e)
            {
                FolderBrowserDialog f = new FolderBrowserDialog();
                f.ShowNewFolderButton = true;
                f.Description = "Selecteer de folder waar de configuratie file opgeslagen moeten worden";
                f.RootFolder = Environment.SpecialFolder.Desktop;
                f.SelectedPath = this.tbConfigFile.Text;
                DialogResult result = f.ShowDialog();
                if (result == DialogResult.OK)
                    this.tbConfigFile.Text = f.SelectedPath;
            }

            private void btnCodeSystemsFile_Click(object sender, EventArgs e)
            {
                FolderBrowserDialog f = new FolderBrowserDialog();
                f.ShowNewFolderButton = true;
                f.Description = "Selecteer de folder waar de 'zib' codestelsels file opgeslagen moeten worden";
                f.RootFolder = Environment.SpecialFolder.Desktop;
                f.SelectedPath = this.tbCodeSystemsFile.Text;
                DialogResult result = f.ShowDialog();
                if (result == DialogResult.OK)
                    this.tbCodeSystemsFile.Text = f.SelectedPath;
            }

            private void tbWikiFiles_TextChanged(object sender, EventArgs e)
            {
                WikiLocationChanged = true;
            }

            private void tbRTFFiles_TextChanged(object sender, EventArgs e)
            {
                RTFLocationChanged = true;
            }

            private void cBNewConfig_CheckedChanged(object sender, EventArgs e)
            {
                if (cbNewConfig.Checked)
                {
                    configChanged = true;
                    tbConfigName.Enabled = true;
                    lbConfigName.Enabled = true;
                }
                else
                {
                    configChanged = false;
                    tbConfigName.Enabled = false;
                    lbConfigName.Enabled = false;
                    setFormTitle(configName);
                }
            }

            private void tbConfigName_TextChanged(object sender, EventArgs e)
            {
                setFormTitle(tbConfigName.Text);
            }

            private void tbBitsPassword_TextChanged(object sender, EventArgs e)
            {
                bitsPasswordChanged = true;
            }

            private void tbWikiPassword_TextChanged(object sender, EventArgs e)
            {
                wikiPasswordChanged = true;
            }

            private void cbTemplate_SelectedIndexChanged(object sender, EventArgs e)
            {
                this.OnTemplateChanged(new EventWithStringArgs(cbTemplate.Text));

            }

            private void cbReleaseName_SelectedIndexChanged(object sender, EventArgs e)
            {
                // 12-12-2025 this.OnReleaseChanged(new EventWithStringArgs(cbReleaseName.Text));
                tbReleaseInfo.Text = parent.releaseList[cbReleaseName.SelectedIndex][1];
                lblADRepository.Text = parent.releaseList[cbReleaseName.SelectedIndex][2];

                tbNumber.Text = parent.releaseList[cbReleaseName.SelectedIndex][4];
                if (int.Parse(parent.releaseList[cbReleaseName.SelectedIndex][4]) == 0)
                {
                    tbNumber.Visible = false;
                    lblNumber.Visible = false;
                    lblReleaseName.Text = "Publicatie:";
                    lblReleaseName.Left = publicationLabelLeft;
                }
                else
                {
                    tbNumber.Visible = true;
                    lblNumber.Visible = true;
                    lblReleaseName.Text = "Pre-publicatie:";
                    lblReleaseName.Left = publicationLabelLeft - 18;
                }
                this.OnReleaseChanged(new EventWithStringArgs(cbReleaseName.Text + (lblNumber.Visible ? ("-" + tbNumber.Text) : "")));

            }

            private void pubLanguage_Changed(object sender, EventArgs e)
            {
                RadioButton senderButton = (RadioButton)sender;
                if (senderButton.Checked)
                {
                    OnLanguageChanged(new EventWithStringArgs(senderButton.Text));
                }
                return;
            }

            private void BtnImages_Click(object sender, EventArgs e)
            {
                FolderBrowserDialog f = new FolderBrowserDialog();
                f.ShowNewFolderButton = true;
                f.Description = "Selecteer de folder waar de afbeeldingen voor o.a. de spreadsheets staan";
                f.RootFolder = Environment.SpecialFolder.Desktop;
                f.SelectedPath = tbImages.Text;
                DialogResult result = f.ShowDialog();
                if (result == DialogResult.OK)
                {
                    this.tbImages.Text = f.SelectedPath;
                    ImageLocationChanged = true;
                }
                else
                    ImageLocationChanged = false;

            }


            private void StartupConfigFile_TextChanged(object sender, EventArgs e)
            {
                if (!ignoreEvent)
                {
                    bool isConfigButton = ((TextBox)sender).Name == "btnConfigFile";

                    //Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)

                    DialogResult answ = MessageBox.Show("Het wijzigen van de deze locatie vereist het opnieuw opstarten van de applicatie\r\n" +
                                    "De applicatie zal daarom afgesloten worden. Doorgaan?", 
                                    "Wijzigen " + (isConfigButton ? "configuratie locatie" : "codestelsel locatie"), 
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                    if (answ == DialogResult.Yes)
                    {
                        if (isConfigButton)
                        {
                            Settings.application.CommonConfigLocation = tbConfigFile.Text == "" ? Settings.pathCommonAppData : tbConfigFile.Text;
                            if (!Directory.Exists(Settings.application.CommonConfigLocation)) Directory.CreateDirectory(Settings.application.CommonConfigLocation);
                        }
                        else
                        {
                            Settings.application.CodeSystemCodesLocation = tbCodeSystemsFile.Text == "" ? 
                                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) : tbCodeSystemsFile.Text;
                            if (!Directory.Exists(Settings.application.CodeSystemCodesLocation)) Directory.CreateDirectory(Settings.application.CodeSystemCodesLocation);
                        }

                        StartupConfigChanged = true;
                        this.Close();
                    }
                    else
                    {
                        ignoreEvent = true;
                        if (isConfigButton)
                            tbConfigFile.Text = Settings.application.CommonConfigLocation;
                        else
                            tbCodeSystemsFile.Text = Settings.application.CodeSystemCodesLocation;
                        ignoreEvent = false;
                        StartupConfigChanged = false;
                    }
                }      
            }

            private void btnRTFFiles_Click(object sender, EventArgs e)
            {
                FolderBrowserDialog f = new FolderBrowserDialog();
                f.ShowNewFolderButton = true;
                f.Description = "Selecteer de folder waar de aangemaakte PDF bestanden opgeslagen moeten worden";
                f.RootFolder = Environment.SpecialFolder.Desktop;
                f.SelectedPath = tbRTFFiles.Text;
                DialogResult result = f.ShowDialog();
                if (result == DialogResult.OK)
                    this.tbRTFFiles.Text = f.SelectedPath;
            }

            private void tbXMLFiles_TextChanged(object sender, EventArgs e)
            {
                XMLLocationChanged = true;
            }

            private void tbXLSFiles_TextChanged(object sender, EventArgs e)
            {
                XLSLocationChanged = true;
            }
            private void tbZibPrefix_TextChanged(object sender, EventArgs e)
            {
                prefixChanged = true;
            }

            private int Scale(int parameter)
            {
                Size defaultBaseSize = new Size(5, 13);
                return (int)Math.Round((parameter * AutoScaleBaseSize.Height * 1.0d) / (defaultBaseSize.Height * 1.0d), 0, MidpointRounding.AwayFromZero);
            }

            /// <summary>
            /// Handmatig schalen van het configuratie dialog tbv andere resoluties dan 96 dpi omdat FOrms de schaling niet geweldig doet.
            /// Als er groupBoxes etc. bijkomen moeten deze ook hier suspend worden. Lastig, maar als dit deel in de designer staat wist MS het
            /// </summary>
            private void ScaleForm()
            {
                int maxLabelLength;
                int rowDistance = Scale(2);
                int rowCount = 0;
                int groupBoxDistance = Scale(4);
                int elementSpacing = Scale(2);
                int groupBoxBottomPadding = Scale(10);

                //   Stop de layout            
                this.gbZIBs.SuspendLayout();
                this.gbBits.SuspendLayout();
                this.gbWiki.SuspendLayout();
                this.gbWikipageNames.SuspendLayout();
                this.gbRTF.SuspendLayout();
                this.gbLocation.SuspendLayout();
                this.ConfigStatus.SuspendLayout();
                this.SuspendLayout();


                // GroupBox Publicatie informatie
                maxLabelLength = gbZIBs.Controls.OfType<Label>().Select(x => x.Width).Max();
                Control[][] labeledTextBoxes = new Control[][]
                {
                    new Control[] { lblCategory, tbCategory },
                    new Control[] { lblReleaseInfo, tbReleaseInfo },
                    new Control[] { lblZibPrefix, tbZibPrefix }
                };
                foreach (Control[] controls in labeledTextBoxes.OrderBy(z=>z[1].Location.Y))
                {
                    controls[1].Width = Scale(controls[1].Width);
                    //controls[1].Height = Scale(controls[1].Height);
                    controls[1].Location = new Point(maxLabelLength + elementSpacing, controls[1].Top + rowCount * rowDistance);
                    controls[0].Location = new Point(controls[0].Location.X, controls[1].Location.Y + Scale(3));
                    rowCount++; 
                }
                tbReleaseInfo.Width = tbNumber.Right - maxLabelLength;
                tbNumber.Width = Scale(tbNumber.Width);
                tbNumber.Height = Scale(tbNumber.Height);
                tbNumber.Location = new Point(tbNumber.Location.X, tbCategory.Location.Y);
                lblNumber.Location = new Point(tbNumber.Left - lblNumber.Width , lblCategory.Location.Y);
                cbReleaseName.Width = Scale(cbReleaseName.Width);
                cbReleaseName.Height = Scale(cbReleaseName.Height);
                cbReleaseName.Location = new Point(lblNumber.Left - cbReleaseName.Width, tbCategory.Location.Y - 1);  
                lblReleaseName.Location = new Point(cbReleaseName.Left - lblReleaseName.Width, tbCategory.Location.Y);
                rbEN.Location = new Point(rbEN.Location.X, tbZibPrefix.Location.Y + 1);
                rbNL.Location = new Point(rbEN.Left - rbNL.Width, tbZibPrefix.Location.Y + 1);
                lblReleaseLanguage.Location = new Point(lblReleaseName.Right - lblReleaseLanguage.Width, lblZibPrefix.Location.Y);
                gbZIBs.Height = tbZibPrefix.Bottom + groupBoxBottomPadding;

                // GroupBox Publicatie informatie

                gbLocation.Location = new Point(gbLocation.Location.X, gbZIBs.Bottom + groupBoxDistance);
                int textBoxEnd = btnExampleFiles.Left - elementSpacing;

                maxLabelLength = gbLocation.Controls.OfType<Label>().Select(x => x.Width).Max();
                labeledTextBoxes = new Control[][]
                {
                    new Control[] { lblExampleFiles, tbExampleFiles, btnExampleFiles },
                    new Control[] { lblXMLFiles, tbXMLFiles, btnXMLFiles },
                    new Control[] { lblWikiFiles, tbWikiFiles, btnWikiFiles },
                    new Control[] { lblXLSFiles, tbXLSFiles, btnXLSFiles },
                    new Control[] { lblRTFFiles, tbRTFFiles, btnRTFFiles },
                    new Control[] { lblImages, tbImages, btnImages },
                    new Control[] { lblConfigFile, tbConfigFile, btnConfigFile },
                    new Control[] { lblCodeSystemsFile, tbCodeSystemsFile, btnCodeSystemsFile }
                };

                int labelHeight = labeledTextBoxes[0][0].Height;
                int textBoxHeight = labeledTextBoxes[0][1].Height;
                int buttonHeight = labeledTextBoxes[0][2].Height;

                rowCount = 0;
                int buttonYPos = btnExampleFiles.Location.Y;
                foreach (Control[] controls in labeledTextBoxes.OrderBy(z => z[1].Location.Y))
                {
                    controls[2].Location = new Point(gbLocation.Width - controls[2].Width - 2, buttonYPos);
                    controls[1].Width = Scale(controls[1].Width);
                    //controls[1].Height = Scale(controls[1].Height);  Hoogte hoeft zo te zien niet geschaald te worden
                    controls[1].Location = new Point(maxLabelLength + elementSpacing, controls[2].Top + (buttonHeight - textBoxHeight)/2);
                    controls[1].Width = controls[2].Left - controls[1].Left - elementSpacing;
                    controls[0].Location = new Point(controls[0].Location.X, controls[2].Top + (buttonHeight - labelHeight)/2);
                    buttonYPos += controls[2].Height + rowDistance;
                }
                gbLocation.Height = labeledTextBoxes[labeledTextBoxes.Count() -1][2].Bottom + groupBoxBottomPadding;

                // Overige GroupBoxes
                gbBits.Location = new Point(gbBits.Location.X, gbLocation.Bottom + groupBoxDistance);
                gbWiki.Location = new Point(gbWiki.Location.X, gbLocation.Bottom + groupBoxDistance);
                gbRTF.Location = new Point(gbRTF.Location.X, gbBits.Bottom + groupBoxDistance);


                // Knoppen
                int oldButtonY = btnCancel.Location.Y;
                btnCancel.Location = new Point(btnCancel.Location.X, gbWiki.Bottom + Scale(12));
                btnOK.Location = new Point(btnOK.Location.X, btnCancel.Location.Y);
                btnCancel.Location = new Point(btnCancel.Location.X, gbWiki.Bottom + Scale(12));

                int checkBoxHeight = cbNewConfig.Height;

                tbConfigName.Location = new Point(btnOK.Left - Scale(20) - tbConfigName.Width, btnCancel.Top + (buttonHeight-textBoxHeight)/2);
                lbConfigName.Location = new Point(tbConfigName.Left - lbConfigName.Width, btnCancel.Top + (buttonHeight - labelHeight) / 2);
                cbNewConfig.Location = new Point(lbConfigName.Left - cbNewConfig.Width, btnCancel.Top + (buttonHeight - checkBoxHeight) / 2);

                this.Height += btnCancel.Location.Y - oldButtonY;

                this.gbZIBs.ResumeLayout(false);
                this.gbZIBs.PerformLayout();
                this.gbBits.ResumeLayout(false);
                this.gbBits.PerformLayout();
                this.gbWiki.ResumeLayout(false);
                this.gbWiki.PerformLayout();
                this.gbWikipageNames.ResumeLayout(false);
                this.gbWikipageNames.PerformLayout();
                this.gbRTF.ResumeLayout(false);
                this.gbRTF.PerformLayout();
                this.gbLocation.ResumeLayout(false);
                this.gbLocation.PerformLayout();
                this.ConfigStatus.ResumeLayout(false);
                this.ConfigStatus.PerformLayout();
                this.ResumeLayout(false);
                this.PerformLayout();


            }

        }
    }
}
