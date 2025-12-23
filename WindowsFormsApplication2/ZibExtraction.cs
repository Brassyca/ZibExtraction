using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using System.Media;
using System.Xml.Linq;
using System.ComponentModel;
using System.Text;
using Zibs.LabelEditor;
using ZibCodeCheck;
using Zibs.Configuration;
using Zibs.Management;
using Zibs.ExtensionClasses;
using Zibs.EAUtilities;
using CodeManagement;
using ZibIdentifierManagement;
using System.Reflection;


/// <summary>
/// Let bij het opnieuw bouwen van de applicatie na wijzigingen in deze solution of in de onderliggende DLL solutions op de volgorde
/// Deze staat beschreven in het Word document Zib_buildorder.docx dat onderdeel is van deze solution (zie Solution Explorer)
/// </summary>

namespace Zibs
{
    namespace ZibExtraction
    {
        public partial class ZibExtraction : Form
        {

            public EA.Repository r;
            public ZIB zib;
            //            public string locationBase = Path.GetTempPath() + @"ZibExtraction\";
            public string locationBase;
            public List<string> templateList;
            public List<string> configList;
            public List<string[]> releaseList;
            public string configName;
            public textLanguage zibPublishLangauge;
            EA.Package selectedPackage;
            Dictionary<string, string> issueDictionary = new Dictionary<string, string>();
            Dictionary<string, string> treeImageDictionary = new Dictionary<string, string>();
            bool tocDone, uploadDone, batch;

            bool eapFileOpen = false;

            public delegate bool zibAction(Dictionary<string, bool> opt, object sender);
            private Dictionary<string, actionData> zibActions = new Dictionary<string, actionData>(); //Lijst met radiobuttons en acties (opgave)
            private List<zibAction> toDoList = new List<zibAction>(); //(todo acties uit knopopgave)
            private List<EA.Package> zibList = new List<EA.Package>();

            private StringBuilder messageLog = new StringBuilder();
            private StringBuilder errorLog = new StringBuilder();
            private int errorCount;
            // haal de verwachte bitness van EA op (x68/x64)
            private string bitness = Properties.Settings.Default.Bitness;

            private Size defaultBaseSize = new Size(5, 13);
            private bool forceDirInput = false;
            private bool extendedLogging = false;


            // ====================
            // Constructor
            // ====================

            public ZibExtraction()
            {
                InitializeComponent();

                string[] cargs = Environment.GetCommandLineArgs();
                if (cargs.Length > 1)
                {
                    if(cargs[1].ToUpper() == "-D") 
                        forceDirInput = true;
                    else if (cargs[1].ToUpper() == "-L") extendedLogging = true;
                }
                this.FormClosing += ZibExtraction_FormClosing;
                //MessageBox.Show(new Form { TopMost = true }, "Attach" + " Bitness: " + (IntPtr.Size * 8).ToString());

                zibActions.Add("Register", new actionData("ccRegister", "Registeer Zib's", registerZIB, false, false));
                zibActions.Add("References", new actionData("ccReferences", "Registreer referenties", registerRefs, false, false));
                zibActions.Add("XML", new actionData("ccXML", "XMI en XML waardelijsten", getZIB_XMLfiles, true, true));
                zibActions.Add("SingleLanguage", new actionData("ccSingleLanguage", "Eéntalig maken", createSingleLanguageZIB, false, true));
                zibActions.Add("Wiki", new actionData("ccWiki", "Wiki pagina's aanmaken", getZIB_sections, false, true));
                zibActions.Add("RTF", new actionData("ccRTF", "PDF/DOCX bestanden maken", getZIB_RTFfile, false, true));
                zibActions.Add("XLS", new actionData("ccXLS", "XLSX bestanden aanmaken", createZIB_XLSfile, false, true));
                zibActions.Add("WikiTOC", new actionData("ccWikiTOC", "Inhoudsopgave maken", createWikiTOCpage, false, true));
                zibActions.Add("WikiUpload", new actionData("ccWikiUpload", "Wiki pagina's uploaden", uploadWikiFiles, false, true));
                ScaleForm();
                batch = false;
                placeButtonBoxes(batch);
                if (batch) cbBatch.Checked = true;
            }

            private void placeButtonBoxes(bool batch)
            {
                int yStep = 20;
                int xStep = 178;
                int xPos = 4;
                int yPos = 20;
                int i = -1;
                int j, k;
                Control _choiceControl;
                bool RadioButtonChecked = false;

                this.gBAction.SuspendLayout();

                // Scaling voor andere resoluties dan 96 DPI
                if (AutoScaleBaseSize != defaultBaseSize)
                {
                    //                    yStep = (int)Math.Round((yStep * AutoScaleBaseSize.Height * 1.0d) / (defaultBaseSize.Height * 1.0d), 0, MidpointRounding.AwayFromZero);
                    //                    xStep = (int)Math.Round((xStep * AutoScaleBaseSize.Width * 1.0d) / (defaultBaseSize.Width * 1.0d), 0, MidpointRounding.AwayFromZero);
                    yStep = Scale(yStep);
                    xStep = Scale(xStep);
                }
 

                foreach (var ccName in zibActions)
                {
                    if (batch)
                        _choiceControl = new CheckBox();
                    else
                        _choiceControl = new RadioButton();
                    i++;
                    j = i % 5;
                    k = i /5;

                    _choiceControl.AutoSize = true;
                    _choiceControl.Location = new Point(xPos + k  * xStep, (batch? yPos : yPos-1) + j * yStep);
                    _choiceControl.Name = ccName.Value.ctrlName;
                    _choiceControl.Text = ccName.Value.description;
                    if (_choiceControl.GetType().Equals(typeof(RadioButton)))
                    {
                        ((RadioButton)_choiceControl).Checked = ccName.Value.isChecked && !RadioButtonChecked;
                        if (ccName.Value.isChecked) RadioButtonChecked = true;
                        ((RadioButton)_choiceControl).UseVisualStyleBackColor = true;
                        ((RadioButton)_choiceControl).CheckedChanged += new EventHandler(this.gBAction_CheckedChanged);
                        ((RadioButton)_choiceControl).Enabled = true;
                    }
                    else if (_choiceControl.GetType().Equals(typeof(CheckBox)))
                    {
                        ((CheckBox)_choiceControl).Checked = ccName.Value.isChecked && ccName.Value.isBatchable;
                        ((CheckBox)_choiceControl).UseVisualStyleBackColor = true;
                        ((CheckBox)_choiceControl).CheckedChanged += new EventHandler(this.gBAction_CheckedChanged);
                        ((CheckBox)_choiceControl).Enabled = ccName.Value.isBatchable;
                    }
                    gBAction.Controls.Add(_choiceControl);
                }
                if (!batch && !RadioButtonChecked)
                {
                    gBAction.Controls.OfType<RadioButton>().FirstOrDefault().Checked = true;
                }

                this.gBAction.ResumeLayout(false);
                this.gBAction.PerformLayout();
            } 


            private void removeButtonBoxes()
            { 
                foreach (var ccName in zibActions)
                {
                    if (gBAction.Controls.OfType<ButtonBase>().Where(b => b.Name == ccName.Value.ctrlName).Count() > 0)
                        {
                            var buttonBox =gBAction.Controls.OfType<ButtonBase>().Where(b => b.Name == ccName.Value.ctrlName).Single();
                            buttonBox.GetType().GetEvent("CheckedChanged").RemoveEventHandler(buttonBox, new EventHandler(gBAction_CheckedChanged));
                            gBAction.Controls.Remove(buttonBox);
                            buttonBox.Dispose();
                        }
                }
            }

            // ========================
            // Form event handlers
            // ========================

            private void Form1_Load(object sender, EventArgs e)
            {
                string installCheck = "";
                bool doExit = false;

                // 6-9-23 Check de aanwezigheid van de startconfiguratie file en de registry keys
                EAUtils.SetCheckData(Settings.pathCommonAppData, Settings.pathAppData, new List<(string, bool)>() { (Settings.startConfigFileName, false) });
                List<string> result = EAUtils.CheckInstall(bitness, EAUtils.CheckAction.Files).ToList<string>(); ;

                if (!Settings.readStartConfig())
                {
                    MessageBox.Show("Applicatieconfiguratie bestand niet gevonden\r\n" +
                        Path.Combine(Settings.pathAppData, Settings.startConfigFileName) + "\r\n" +
                        "Applicatie kan niet gestart worden", "Foutmelding", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.Close();
                }
                else 
                    // Schrijf de StartConfig meteen weer weg. Hiermee kunnen nieuwe parameters die met default gemaakt zijn opgeslagen worden.
                    // Deze kunnen dan later in de xml file handmatig aangepast worden
                    Settings.writeStartConfig();


                /// Voorkom kip en ei problemen door deze locaties eerst te vullen als ze leeg zijn
                if (string.IsNullOrWhiteSpace(Settings.userPreferences.ImageLocation) ||
                    !Directory.Exists(Settings.userPreferences.ImageLocation) ||
                    string.IsNullOrWhiteSpace(Settings.application.CodeSystemCodesLocation) ||
                    !Directory.Exists(Settings.application.CodeSystemCodesLocation) ||
                    string.IsNullOrWhiteSpace(Settings.application.CommonConfigLocation) ||
                    !Directory.Exists(Settings.application.CommonConfigLocation) ||
                    forceDirInput)
                {
                    frmFileSelection fileSelection = new frmFileSelection(forceDirInput);
                    DialogResult answ = fileSelection.ShowDialog();
                    if (answ == DialogResult.OK)
                    {
                        Settings.writeStartConfig();
                        doExit = false;
                    }

                    else
                        doExit = true;
                    fileSelection.Dispose();
                    if (doExit) Application.Exit();
                }

                // 6-9-23: install check file copy tbv beheerde Pc's
                EAUtils.SetCheckData(Settings.pathCommonAppData, Settings.pathAppData, new List<(string, bool)>() { (Settings.configFileName, true) });
                result.AddRange(EAUtils.CheckInstall(bitness, EAUtils.CheckAction.Both).ToList<string>());
                installCheck += "** Installatiecheck resultaat =\r\n" + string.Join("\r\n", result.ToArray()) + "\r\n** Einde check resultaat";

                string exceptionText = "";
                EAUtils.BaseUrl = Settings.application.ManualWikiLocation;

                int supportedVersionTest = EAUtils.SupportedVersion(bitness, out Version currentVersion, out string messages);

                if (extendedLogging)
                {
                    MessageBox.Show(installCheck + "\r\n" + messages +
                        "\r\nVerwachte bitNess: " + bitness +
                        "\r\nApplicatie bitNess: " + (Environment.Is64BitProcess ? "64" : "32") +
                        "\r\nOS bitNess: " + (Environment.Is64BitOperatingSystem ? "64" : "32") +
                        "\r\nValidatie resultaat= " + supportedVersionTest, "Start log"); ;
                }
                if (supportedVersionTest < 0)
                {
                    switch (supportedVersionTest)
                    {
                        case -1:
                            exceptionText = "EA version not validated";
                            break;
                        case -2:
                            exceptionText = "License expired";
                            break;
                        case -3:
                            exceptionText = "General application error";
                            break;
                        case -4:
                            exceptionText = "License tampered";
                            break;
                        case -5:
                            exceptionText = "Is Enterprise Architect installed?";
                            break;
                        case -6:
                            exceptionText = "Wrong ZibExtraction bitness version for 64bits/32bit version of EA";
                            break;
                        case -7:
                            exceptionText = "Checkdata not accessible";
                            break;
                        case -8:
                            exceptionText = "Registry keys not installed";
                            break;
                        default:
                            exceptionText = "Undocumented exception";
                            break;
                    }
                    //throw (new Exception(exceptionText));
                    MessageBox.Show(exceptionText +
                             "\r\nHet programma wordt afgesloten", "Unhandled exception", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    Application.Exit();
                }

                // Version nag screen
                Version thisVersion = Assembly.GetExecutingAssembly().GetName().Version;
                if (currentVersion  != null && thisVersion < currentVersion)
                {
                    MessageBox.Show(" Dit is niet de nieuwste versie van ZibExtraction.\r\n Huidige versie: " + string.Format("{0}.{1}", thisVersion.Major, thisVersion.Minor) + ", de nieuwste versie: " + currentVersion.ToString()
                        + "\r\n\r\nUpgrade naar de nieuwste versie om deze mededeling niet meer te tonen.", "Nieuwe versie beschikbaar", MessageBoxButtons.OK ,MessageBoxIcon.Information);
                }

                // set de encryptie delegate
                Settings.setEncryption();

                Settings.EA_Version = (string)Microsoft.Win32.Registry.GetValue("HKEY_CURRENT_USER\\Software\\Sparx Systems\\" + (bitness.Contains("64") ? "EA64" : "EA400") + "\\EA", "Version", string.Empty);

                locationBase = Settings.application.WorkLocation + "ZibExtraction\\";
                Settings.readGroupAndReleaseManagementFile(out releaseList);
                this.Text = "ZIB Extractor (" + bitness + "bits EA)";



                // vul de treeview imagesList
                // de closed versie moet voor de open versie staan

                ImageList languageImages = new ImageList();
                languageImages.Images.Add("Empty", Properties.Resources.Dash1);
                languageImages.Images.Add("NL", Properties.Resources.nl_icon_20px);
                languageImages.Images.Add("EN", Properties.Resources.gb_icon_20px);
                languageImages.Images.Add("Question", Properties.Resources.Help_icon_20px);
                languageImages.Images.Add("FolderClosed", Properties.Resources.map_c);
                languageImages.Images.Add("FolderOpen", Properties.Resources.map_o);
                projectView.ImageList = languageImages;
                projectView.ImageIndex = 0;
                projectView.SelectedImageIndex = 0;
                treeImageDictionary.Add("other", "FolderClosed");
                treeImageDictionary.Add("zibcontainer", "FolderClosed");
                treeImageDictionary.Add("singlezib", "Question");
                treeImageDictionary.Add("NL", "NL");
                treeImageDictionary.Add("EN", "EN");
                treeImageDictionary.Add("Multi", "Question");
                gBAction_CheckedChanged(sender, e);

                /*
                if (gBAction.Controls.OfType<RadioButton>().Where(c => c.Name == "ccRegister").Count() > 0)
                {
                    (gBAction.Controls.OfType<RadioButton>().Where(c => c.Name == "ccRegister").Single()).PerformClick();
                    gBAction_CheckedChanged(sender, e);

                } 
                
                if (gBAction.Controls.OfType<RadioButton>().Count() > 0)
                {
                    (gBAction.Controls.OfType<RadioButton>().Where(c => c.Checked).First()).PerformClick();
                    gBAction_CheckedChanged(sender, e);

                }
                if (gBAction.Controls.OfType<CheckBox>().Count() > 0)
                {
                    var checkedBoxes = gBAction.Controls.OfType<CheckBox>().Where(c => c.Checked);
                    foreach (CheckBox cb in checkedBoxes)
                    {
                        cb.CheckState = CheckState.Checked;
                        gBAction_CheckedChanged(sender, e);
                    }   
                }
                */

                dudLogLevel.Items.AddRange(new string[] { "Information", "Warning", "Error", "Disaster"});
                dudLogLevel.SelectedItem = "Information";
/*

                // resize voor het geval er een andere resolutie dan 96 dpi 
                gbSettings.Location = new Point(gbSettings.Location.X, gbExecute.Location.Y + gbExecute.Size.Height + 3);
                if (gbOption.Size.Height > gBAction.Size.Height) gbOption.Size = new Size(gbOption.Size.Width, gBAction.Size.Height);

                int newSplitterDistance = splitContainer2.Height - (gbSettings.Location.Y + gbSettings.Size.Height +15);
                splitContainer2.SplitterDistance = newSplitterDistance;
                gbResult.Location = new Point(gbExecute.Location.X + gbExecute.Width + 6, gbResult.Location.Y);
                int endOptionBox = gbOption.Location.X + gbOption.Size.Width;
                int endResultBox = gbResult.Location.X + gbResult.Size.Width;
                int endSettingsBox = gbSettings.Location.X + gbSettings.Size.Width;
                int shortestBox = Math.Min(endOptionBox, Math.Min(endResultBox, endSettingsBox));
                int largestBox = Math.Max(endOptionBox, Math.Max(endResultBox, endSettingsBox));
                if (shortestBox < gbResult.Location.X + btResultLocation.Location.X + btResultLocation.Width)
                {
                    shortestBox = gbResult.Location.X + btResultLocation.Location.X + btResultLocation.Width;
                    // voer de nieuwe size door en kijk of winforms deze aanpast.
                    gbResult.Size = new Size(shortestBox - gbResult.Location.X, gbResult.Size.Height);
                    shortestBox = gbResult.Location.X + gbResult.Size.Width;
                }
                gbOption.Size = new Size(shortestBox - gbOption.Location.X, gbOption.Size.Height);
                gbResult.Size = new Size(shortestBox - gbResult.Location.X, gbResult.Size.Height);
                gbSettings.Size = new Size(shortestBox - gbSettings.Location.X, gbSettings.Size.Height);
                this.Width = splitContainer1.Panel1.Width + splitContainer1.SplitterWidth + shortestBox + 40;
*/
            }


            /// <summary>
            /// TOOPSTRIP MENU HANDLERS
            /// </summary>

            private void openToolStripMenuItem_Click(object sender, EventArgs e)
            {
                OpenFileDialog d = new OpenFileDialog();
                string filter = (bitness.Contains("32") ? "*.eap; *.eapx" : "*.qea;*.qeax") + "; *.easvrlnk";
                d.Filter = "Enterprise Architect files (" + filter + ")|" + filter + "|All files (*.*)|*.*";
                d.FilterIndex = 1;
                d.Title = "Open EA file";

                if (d.ShowDialog() == DialogResult.OK)
                {
                    this.lblFileNameValue.Text = d.FileName.EllipsisFilename(lblFileNameValue.MaximumSize, this.lblFileNameValue.Font);
                    OpenFile(d.FileName);

                }
            }


            private void ZibExtraction_FormClosing(object sender, FormClosingEventArgs e)
            {
                if (this.zibExtractor.IsBusy)
                {
                    this.StatusLabel.Text = "Waiting for process to finish...";
                    EAStatus.Refresh();
                    zibExtractor.RunWorkerCompleted += new RunWorkerCompletedEventHandler((sender3, e3) => this.Close());
                    if (!this.zibExtractor.CancellationPending) this.zibExtractor.CancelAsync();
                    e.Cancel = true;
                    return;
                }
                this.StatusLabel.Text = "Closing...";
                EAStatus.Refresh();
                try
                {
                    if (!(r == null))
                    {
                        r.CloseFile();
                        r.Exit();
                        r = null;
                        eapFileOpen = false;
                    }
                    deleteTempFiles();
                    deleteEmptyDirectories();
//                    Properties.Settings.Default.CommonConfigLocation = Settings.application.CommonConfigLocation;
//                    Properties.Settings.Default.Save();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Fout " + ex.ToString());
                }
            }


            private void sluitenToolStripMenuItem_Click(object sender, EventArgs e)
            {
                if (this.zibExtractor.IsBusy)
                {
                    zibExtractor.RunWorkerCompleted += zibExtractor_Closing;
                    this.StatusLabel.Text = "Waiting for process te finish...";
                    EAStatus.Refresh();
                    if (!this.zibExtractor.CancellationPending) this.zibExtractor.CancelAsync();
                    return;
                }

                zibExtractor.RunWorkerCompleted -= zibExtractor_Closing;

                this.StatusLabel.Text = "Closing...";
                EAStatus.Refresh();

                r.CloseFile();
                if (!(r == null)) 
                {
                    r.Exit();
                    r = null;
                    eapFileOpen = false;
                }
                selectedPackage = null;
                clearForm();
                deleteTempFiles();
                deleteEmptyDirectories();
                messageLog.Clear();
                errorLog.Clear();
                this.openToolStripMenuItem.Enabled = true;
                this.configuratieToolStripMenuItem.Enabled = false;
                this.btnActie.Enabled = false;
                this.StatusLabel.Text = "EAP file closed";
                EAStatus.Refresh();
            }


            private void zibExtractor_Closing(object sender, RunWorkerCompletedEventArgs e)
            {
                this.sluitenToolStripMenuItem_Click(sender, e);
            }


            private void exitToolStripMenuItem_Click(object sender, EventArgs e)
            {
                this.Close();
            }

            private void publicatiesEnGroepenToolStripMenuItem_Click(object sender, EventArgs e)
            {
                string fileName = Path.Combine(Settings.application.CommonConfigLocation, Settings.zibRegistryFileName);
                frmManagement mgntConfig = new frmManagement(fileName);
                mgntConfig.ShowDialog();
                // Lees de management file opnieuw uit om wijzigingen te effectueren
                Settings.readGroupAndReleaseManagementFile(out releaseList);
            }

            private void configuratieToolStripMenuItem_Click(object sender, EventArgs e)
            {
                frmConfig config = new frmConfig(this.lblConfigValue.Text, this);
                config.LanguageChanged += Config_LanguageChanged;
                config.PrefixChanged += Config_PrefixChanged;
                config.TemplateChanged += Config_TemplateChanged;
                config.ReleaseChanged += Config_ReleaseChanged;
                config.ConfigChanged += Config_ConfigChanged;
                config.ShowDialog();
                Settings.saveToFile(configName, locationBase);
                if (config.StartupConfigChanged || config.ImageLocationChanged)
                {
                    if (!Settings.writeStartConfig())
                        MessageBox.Show("Applicatieconfiguratie bestand niet gevonden\r\n" +
                            Path.Combine(Settings.pathAppData, Settings.startConfigFileName) + "\r\n" +
                            "Wijzigingen kunnen niet opgeslagen worden", "Foutmelding", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (config.StartupConfigChanged)
                    {
                        config.Dispose();
                        this.Close();
                    }
                }
                this.lblWikiSiteValue.Text = Settings.wikicontext.wikiBaserurl;
                this.lblExampleDirValue.Text = Settings.userPreferences.ExampleLocation.EllipsisFilename(lblExampleDirValue.MaximumSize, lblExampleDirValue.Font);
                this.lblPrefixValue.Text = Settings.zibcontext.zibPrefix;
                config.Dispose();
            }

            private void taallabelsToolStripMenuItem_Click(object sender, EventArgs e)
            {
                frmLabelEditor labelEdit = new frmLabelEditor();
                labelEdit.FileName = Path.Combine(Settings.application.CommonConfigLocation, Settings.zibLabelFileName);
                labelEdit.ShowDialog();
                labelEdit.Dispose();

            }

            private void codecheckToolStripMenuItem_Click(object sender, EventArgs e)
            {
                frmMain codeCheck = new frmMain();
                codeCheck.ShowDialog();
                codeCheck.Dispose();
            }


            private void zibCodestelselBeheerToolStripMenuItem_Click(object sender, EventArgs e)
            {
                CodeManager codemanager = new CodeManager();
                codemanager.ShowDialog();
                codemanager.Dispose();
            }

            private void zibIdBeheerToolStripMenuItem_Click(object sender, EventArgs e)
            {
                ZibIdentifierManager zibIdentifierManager = new ZibIdentifierManager();
                zibIdentifierManager.ShowDialog();
                zibIdentifierManager.Dispose();
            }




            private void configuratiesToolStripMenuItem_Click(object sender, EventArgs e)
            {
                frmManageConfigurations configManager = new frmManageConfigurations(configName);
                configManager.ShowDialog();
                configManager.Dispose();
            }


            private void overZibExtractionToolStripMenuItem_Click(object sender, EventArgs e)
            {
                this.about(sender, e);
            }

            private void helppaginaToolStripMenuItem_Click(object sender, EventArgs e)
            {
                System.Diagnostics.Process.Start(Path.Combine(Settings.application.ManualWikiLocation + "ZibExtraction"));
            }


            /// <summary>
            /// CONTEXT MENU HANDLERS VOOR HET RESULT WINDOW
            /// </summary>

            private void clearToolStripMenuItem_Click(object sender, EventArgs e)
            {
                Result.Clear();
            }
            private void copyToolStripMenuItem_Click(object sender, EventArgs e)
            {
                Result.SelectAll();
                Result.Copy();
            }


            /// <summary>
            /// PROJECTVIEW SELECTIE EVENT HANDLERS
            /// </summary>

            private void projectView_AfterSelect(object sender, TreeViewEventArgs e)
            {
                //            clearElementWindows();

                selectedPackage = (EA.Package)e.Node.Tag;
                if (packageType(selectedPackage) != PackageType.other)
                {
                    this.lblSelectedValue.Text = selectedPackage.Name;
                    this.lblZibLanguageValue.Text = packageType(selectedPackage) == PackageType.singlezib ? (selectedPackage.Element.TaggedValuesEx.GetByName("HCIM::PublicationLanguage")?.Value ?? "Multi") : "n.a.";

                    if(!zibExtractor.IsBusy) this.btnActie.Enabled = true; //als het achtergrondproces een repaint heeft gevraagd, mag de knop niet vrijgegeven worden
                }
                else
                {
                    this.lblSelectedValue.Text = "Geen geldig package";
                    this.lblZibLanguageValue.Text = "";
                    SystemSounds.Exclamation.Play();
                    this.btnActie.Enabled = false;
                }
            }

            private void projectView_AfterExpand(object sender, TreeViewEventArgs e)
            {
                string key = (string)e.Node.ImageKey.ToString().Replace("Closed", "Open");
                if (key != "-1") e.Node.ImageKey = key;
            }

            private void projectView_AfterCollapse(object sender, TreeViewEventArgs e)
            {
                string key = (string)e.Node.ImageKey.ToString().Replace("Open", "Closed");
                if (key != "-1") e.Node.ImageKey = key;
            }

            // =======================
            // CHECKBOX EVENT HANDLERS
            // =======================

            private void gBAction_CheckedChanged(object sender, EventArgs e)
            {
                //                Type type = sender.GetType();
                bool anyBoxCheck = false;
                for (int i = 0; i < gbOption.Controls.Count; i++)
                    gbOption.Controls[i].Enabled = false;
                dudLogLevel.Enabled = true;
                lbLogLevel.Enabled = true;
                cbSaveTempDirs.Enabled = true;
                cBOldFormat.Enabled = true;
                // check of er minimaal één checkbox/radiobutton gecheckt is
                anyBoxCheck = gBAction.Controls.OfType<Control>().Where(x => (bool)x.GetType().GetProperty("Checked").GetValue(x, null)).Any();
                bool wikiUploadOnly = anyBoxCheck;
                bool wikiTOCOnly = anyBoxCheck;

                foreach (Control buttonBox in gBAction.Controls)
                    if (buttonBox.GetType().Equals(typeof(RadioButton)) || buttonBox.GetType().Equals(typeof(CheckBox)))
                        if ((bool)buttonBox.GetType().GetProperty("Checked").GetValue(buttonBox, null))        // is de checkbox of radiobutton checked?
                        {
                            if (buttonBox.Name == "ccWikiUpload")
                                wikiUploadOnly &= true;
                            else
                                wikiUploadOnly &= false;
                            if (buttonBox.Name == "ccWikiTOC")
                                wikiTOCOnly &= true;
                            else
                                wikiTOCOnly &= false;
                            if (buttonBox.Name == "ccWiki")
                                cBwikiPreview.Enabled = true;
                            if (buttonBox.Name == "ccWikiUpload")
                            {
                                cBNoCertificateErrors.Enabled = true;
                                cBPurgePages.Enabled = true;
                            }
                            if (buttonBox.Name == "ccXLS")
                                cBwikiPreview.Enabled = true;
                            if (buttonBox.Name == "ccXML")
                                cBArtDecor.Enabled = true;
                            if (buttonBox.Name == "ccRegister")
                            {
                                cBForceReg.Enabled = true;
                                cbBackup.Enabled = true;
                            }
                            if (buttonBox.Name == "ccReferences")
                            { 
                                cBForceReg.Enabled = true;
                                cbBackup.Enabled = true;
                            }
                            if (buttonBox.Name == "ccWikiTOC")
                                cBwikiPreview.Enabled = true;
                            if (buttonBox.Name == "ccRTF")
                                cBDocx.Enabled = true;
                            // if (buttonBox.Name == "ccSingleLanguage")
                            // cBwikiPreview.Enabled = true;
                        }

                bool aa = batch && !gBAction.Controls.OfType<CheckBox>().Where(x => x.Checked).Any();
                if (batch && !gBAction.Controls.OfType<CheckBox>().Where(x => x.Checked).Any())
                {
                    btnActie.Enabled = false;
                }
                else
                {
                    if(!(selectedPackage == null || packageType(selectedPackage) == PackageType.other))
                        btnActie.Enabled = true;
                    else
                        btnActie.Enabled = false;
                }


                if (!wikiUploadOnly)
                {
                    cbReuse.Checked = false;
                    cbReuse.Enabled = false;
                }
                else
                {
                    cbReuse.Enabled = true;
                }

                if (wikiTOCOnly || wikiUploadOnly)
                {
                    if (eapFileOpen) btnActie.Enabled = true;
                }
                else
                {
                    if (selectedPackage == null || packageType(selectedPackage) == PackageType.other) btnActie.Enabled = false;
                }

            }

            private void cBNoCertificateErrors_CheckedChanged(object sender, EventArgs e)
            {
                if (cBNoCertificateErrors.Checked)
                    Settings.ignoreCertificateErrors = true;
                else
                    Settings.ignoreCertificateErrors = false;
            }


            private void cbForceReg_CheckedChanged(object sender, EventArgs e)
            {
                if (cBForceReg.Checked)
                    Settings.forceZibRegistration = true;
                else
                    Settings.forceZibRegistration = false;
            }

            private void cBOldFormat_CheckedChanged(object sender, EventArgs e)
            {
                if (cBOldFormat.Checked)
                    Settings.oldFormat = true;
                else
                    Settings.oldFormat = false;
            }

            private void cbBatch_CheckedChanged(object sender, EventArgs e)
            {
                string firstChecked;
                if (cbBatch.Checked)
                    batch = true;
                else
                    batch = false;

                var allChecked = gBAction.Controls.OfType<ButtonBase>().Where(x => ((bool)x.GetType().GetProperty("Checked").GetValue(x, null)) == true);
                if (allChecked.Count() > 0)
                    firstChecked = allChecked.First().Name.Substring(2);
                else
                    firstChecked = zibActions.Keys.First();

                foreach (var actionName in  zibActions.Keys)
                    zibActions[actionName].isChecked = false;
                zibActions[firstChecked].isChecked = true;

                removeButtonBoxes();
                placeButtonBoxes(batch);
                // zet de opties goed
                gBAction_CheckedChanged(sender, e);

                var languageButtons = gBAction.Controls.OfType<ButtonBase>().Where(x => x.Name == "ccSingleLanguage");
                if (languageButtons.Count() > 0)
                    languageButtons.First().Text = zibActions["SingleLanguage"].description + " (" + Settings.zibcontext.pubLanguage.ToString() + ")";



            }
            private void cBDocx_CheckedChanged(object sender, EventArgs e)
            {
                if (cBDocx.Checked)
                    Settings.PDFFormat = false;
                else
                    Settings.PDFFormat = true;
            }

            private void cbSave_CheckedChanged(object sender, EventArgs e)
            {
                if (cbSave.Checked)
                {
                    cbReuse.Checked = false;
                    tbResultLocation.ReadOnly = false;
                    btResultLocation.Enabled = true;
                    if (tbResultLocation.Text == "") btResultLocation.PerformClick();

                }
                else
                {
                    tbResultLocation.ReadOnly = true;
                    btResultLocation.Enabled = false;
                }
            }


            /// <summary>
            /// Reuse checkbox summary: geeft aan of eerder opgeslagen resultaten hergebruikt worden ipv dat nieuwe 
            /// resultaten gegenereerd worden.
            /// Wordt gebruikt om eerder voor de testomgeving aangemaakte resultaten, na testen, te hergebruiken voor
            /// de produktie omgeving.
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            private void cbReuse_CheckedChanged(object sender, EventArgs e)
            {
                if (cbReuse.Checked)
                {
                    cbSave.Checked = false;
                    tbResultLocation.ReadOnly = false;
                    btResultLocation.Enabled = true;
                    // Uncheck alle buttons behalve WikiUpload
                    foreach (Control buttonBox in gBAction.Controls)
                    {
                        if (buttonBox.GetType().Equals(typeof(RadioButton)) || buttonBox.GetType().Equals(typeof(CheckBox)))
                            if (buttonBox.Name == "ccWikiUpload")
                                buttonBox.GetType().GetProperty("Checked").SetValue(buttonBox, true);
                            else
                                buttonBox.GetType().GetProperty("Checked").SetValue(buttonBox, false);
                    }
                    if (tbResultLocation.Text == "") btResultLocation.PerformClick();
                }
                else
                {
                    tbResultLocation.ReadOnly = true;
                    btResultLocation.Enabled = false;
                }
            }

            private void cBPurgePages_CheckedChanged(object sender, EventArgs e)
            {

                if (cBPurgePages.Checked)
                    Settings.doPurgePages = true;
                else
                    Settings.doPurgePages = false;
            }


            /// <summary>
            /// BUTTON EVENT HANDLERS
            /// </summary>

            /// <summary>
            /// Action button event handler
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            ///
            private void btnActie_Click(object sender, EventArgs e)
            {
                tocDone = false;
                uploadDone = false;
                zibList.Clear();
                toDoList.Clear();
                if (packageType(selectedPackage) == PackageType.singlezib)
                    zibList.Add(selectedPackage);
                else if (packageType(selectedPackage) == PackageType.zibcontainer)
                {
                    if (projectView.SelectedNode.Nodes.Count == 0)
                    {
                        SystemSounds.Exclamation.Play();
                        return;
                    }
                    DialogResult answ = MessageBox.Show("Actie op alle zib's in dit package", "Waarschuwing"
                   , MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (answ == DialogResult.OK)
                    {
                        if (selectedPackage.Packages.Count > 0)
                            foreach (EA.Package child in selectedPackage.Packages)
                                if (child.Name.Contains(Settings.zibcontext.zibPrefix))
                                    zibList.Add(child);
                    }
                    else return;
                }
                else
                {
                    SystemSounds.Exclamation.Play();
                    return;
                }

                foreach (Control buttonBox in gBAction.Controls)
                    if (buttonBox.GetType().Equals(typeof(RadioButton)) || buttonBox.GetType().Equals(typeof(CheckBox)))
                        if ((bool)buttonBox.GetType().GetProperty("Checked").GetValue(buttonBox, null))        // is de checkbox of radiobutton checked?
                            toDoList.Add(zibActions[buttonBox.Name.Substring(2)].function);

                if (batch)
                {
                    var cbCollection = gBAction.Controls.OfType<CheckBox>();
                    int checkCount = cbCollection.Where(c => c.Checked).Count();

                    if (!((cbCollection.Where(c => c.Name == "ccSingleLanguage").Single().Checked) ||
                            (checkCount == 1 && cbCollection.Where(c => c.Name == "ccXML").Single().Checked)))
                    {
                        DialogResult answ = MessageBox.Show("De gekozen acties kunnen alleen op ééntalige zib's uitgevoerd worden. Doorgaan?", "Waarschuwing"
                                , MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (answ == DialogResult.No) return;
                    }
                }
                zibExtractorPars zp = new zibExtractorPars();
                zp.actionList = toDoList;
                zp.options = gbOption.Controls.OfType<CheckBox>().ToDictionary(x => x.Name, x => x.Checked);
                foreach (var kvp in gbResult.Controls.OfType<CheckBox>().ToDictionary(x => x.Name, x => x.Checked))
                    zp.options.Add(kvp.Key, kvp.Value);
                zp.copyPath = tbResultLocation.Text;
                bool levelOk = ErrorType.TryParse(dudLogLevel.Text.ToLower(), out ErrorType errorLevel);
                zp.logLevel = levelOk? errorLevel: ErrorType.information;
                btCancel.Enabled = true;
                btnActie.Enabled = false;

                zibExtractor.RunWorkerAsync(zp);
            }
            /// <summary>
            ///  ResultLocation button event handler
            /// </summary>

            private void btResultLocation_Click(object sender, EventArgs e)
            {
                FolderBrowserDialog f = new FolderBrowserDialog();
                f.ShowNewFolderButton = true;
                if (cbReuse.Checked)
                    f.Description = "Selecteer de folder waar de resutaten staan";
                else
                    f.Description = "Selecteer de folder waar de resutaten opgeslagen moeten worden";
                f.RootFolder = Environment.SpecialFolder.Desktop;
                f.SelectedPath = tbResultLocation.Text;
                DialogResult result = f.ShowDialog();
                if (result == DialogResult.OK)
                    this.tbResultLocation.Text = f.SelectedPath;
            }


            private void btCancel_Click(object sender, EventArgs e)
            {
                {
                    // Cancel the asynchronous operation.
                    this.zibExtractor.CancelAsync();
                    // Disable the Cancel button.
                    btCancel.Enabled = false;
                    this.StatusLabel.Text = "Cancelling ... ";
                    EAStatus.Refresh();
                }
            }



            /// <summary>
            /// EVENT HANDLERS VOOR AANGEROEPEN CLASSES
            /// </summary>



            private void Config_ConfigChanged(object sender, EventWithStringArgs e)
            {
                this.lblConfigValue.Text = e.Text;
                configName = e.Text;
            }

            private void Config_LanguageChanged(object sender, EventWithStringArgs e)
            {
                var languageButtons = gBAction.Controls.OfType<Control>().Where(x => x.Name == "ccSingleLanguage");
                if (languageButtons.Count() > 0) languageButtons.First().Text = zibActions["SingleLanguage"].description + " (" + e.Text + ")";
                this.lblReleaseLanguageValue.Text = e.Text;
            }
            private void Config_TemplateChanged(object sender, EventWithStringArgs e)
            {
                this.lblTemplateValue.Text = e.Text;
            }

            private void Config_ReleaseChanged(object sender, EventWithStringArgs e)
            {
                this.lblReleaseValue.Text = e.Text;
            }

            private void Config_PrefixChanged(object sender, EventArgs e)
            {
                populateProjectView(saveState: false);
            }

            //
            /// <summary>
            /// PROJECTVIEW CONTROL METHODES
            /// </summary>

            private void populateProjectView(bool saveState)
            {
                Dictionary<string, bool> nodeExpanded = new Dictionary<string, bool>();
                string selectedZibGUID = "";

                TreeNode packageNode;
                if (saveState) nodeExpanded = saveTreeState(projectView, out selectedZibGUID);
                projectView.Nodes.Clear();
                foreach (EA.Package m in r.Models)
                {
                    packageNode = new TreeNode(m.Name);
                    packageNode.Tag = m;
                    packageNode.ImageKey = treeImageDictionary[packageType(m).ToString()];
                    packageNode.SelectedImageKey = treeImageDictionary[packageType(m).ToString()];
                    projectView.Nodes.Add(packageNode);
                    int i = projectView.GetNodeCount(false) - 1;
                    addPackages(m, projectView.Nodes[i].Nodes);
                }
                if (saveState) restoreTreeState(projectView, nodeExpanded, selectedZibGUID);
            }

            Dictionary<string, bool> saveTreeState(TreeView tree, out string selectedTagId)
            {
                Dictionary<string, bool> nodeStates = new Dictionary<string, bool>();
                for (int i = 0; i < tree.Nodes.Count; i++)
                {
                    if (tree.Nodes[i].Nodes.Count > 0)
                    {
                        nodeStates.Add(tree.Nodes[i].Text, tree.Nodes[i].IsExpanded);
                        saveBrancheState(tree.Nodes[i], ref nodeStates);
                    }
                }
                selectedTagId = ((EA.Package)tree.SelectedNode.Tag).PackageGUID.ToString();
                return nodeStates;
            }

            void saveBrancheState(TreeNode branch, ref Dictionary<string, bool> nodeStates)
            {
                for (int i = 0; i < branch.Nodes.Count; i++)
                {
                    if (branch.Nodes[i].Nodes.Count > 0)
                    {
                        nodeStates.Add(branch.Nodes[i].Text, branch.Nodes[i].IsExpanded);
                    }
                    //              recurse
                    saveBrancheState(branch.Nodes[i], ref nodeStates);
                }

                return;
            }


            void restoreTreeState(TreeView tree, Dictionary<string, bool> nodeStates, string selectedTagId)
            {
                TreeNode targetNode = null;
                for (int i = 0; i < tree.Nodes.Count; i++)
                {
                    if (nodeStates.ContainsKey(tree.Nodes[i].Text))
                    {
                        if (nodeStates[tree.Nodes[i].Text])
                            tree.Nodes[i].Expand();
                        else
                            tree.Nodes[i].Collapse();
                    }
                    restoreBranchState(tree.Nodes[i], nodeStates, selectedTagId, ref targetNode);
                }
                var nodes = tree.Nodes.OfType<TreeNode>().Where(n => ((EA.Package)n.Tag).PackageGUID.ToString() == selectedTagId);
                if (nodes.Count() == 1) targetNode = nodes.First();
                tree.SelectedNode = targetNode;
                return;
            }

            void restoreBranchState(TreeNode branch, Dictionary<string, bool> nodeStates, string selectedTagId, ref TreeNode targetNode)
            {
                for (int i = 0; i < branch.Nodes.Count; i++)
                {
                    if (nodeStates.ContainsKey(branch.Nodes[i].Text))
                    {
                        if (nodeStates[branch.Nodes[i].Text])
                            branch.Nodes[i].Expand();
                        else
                            branch.Nodes[i].Collapse();
                    }
                    var nodes = branch.Nodes.OfType<TreeNode>().Where(n => ((EA.Package)n.Tag).PackageGUID.ToString() == selectedTagId);
                    if (nodes.Count() == 1) targetNode = nodes.First();
                    //              recurse
                    restoreBranchState(branch.Nodes[i], nodeStates, selectedTagId, ref targetNode);
                }
                return;
            }

            private void addPackages(EA.Package p, TreeNodeCollection myTree)
            {
                TreeNode packageNode;
                foreach (EA.Package cp in p.Packages.OfType<EA.Package>().OrderBy(x=>x.Name))
                {
                    string imageName;
                    packageNode = new TreeNode(cp.Name);
                    packageNode.Tag = cp;
                    if (packageType(cp).ToString() == "singlezib")
                    {
                        imageName = ZIB.getZibPublishLanguage(cp).ToString();
                    }
                    else
                    {
                        imageName = packageType(cp).ToString();
                    }
                    packageNode.ImageKey = treeImageDictionary[imageName];
                    packageNode.SelectedImageKey = treeImageDictionary[imageName];

                    // voeg de packages toe maar bij zib's alleen als ze de juiste prefix hebben.
                    int startSearch = cp.Name.IndexOf("-v") == -1 ? cp.Name.Length - 1 : cp.Name.IndexOf("-v");
                    bool nameValid = cp.Name.LastIndexOf(".", startSearch) == -1? false: cp.Name.Substring(0, cp.Name.LastIndexOf(".", startSearch)) == Settings.zibcontext.zibPrefix;
//                    if (cp.StereotypeEx != "DCM" || (cp.StereotypeEx == "DCM" && cp.Name.Contains(Settings.zibcontext.zibPrefix)))
                    if (cp.StereotypeEx != "DCM" || (cp.StereotypeEx == "DCM" && nameValid))
                    {
                        myTree.Add(packageNode);
                    }
                    int i = myTree.Count - 1;
                    if (cp.StereotypeEx != "DCM") addPackages(cp, myTree[i].Nodes);
                }

            }

            /// <summary>
            /// INITIALISATIE METHODES
            /// </summary>

            private void setPreferences(string baseDirectory)   //  = Settings.userPreferences.sessionBase
            {
                Settings.userPreferences.VS_RTFLocation = baseDirectory + @"rtf\";
                if (string.IsNullOrWhiteSpace(Settings.userPreferences.XMLLocation)) Settings.userPreferences.XMLLocation = getStandardDirectory("XML", baseDirectory);
                if (string.IsNullOrWhiteSpace(Settings.userPreferences.XLSLocation)) Settings.userPreferences.XLSLocation = getStandardDirectory("XLS", baseDirectory);
                if (string.IsNullOrWhiteSpace(Settings.userPreferences.WikiLocation)) Settings.userPreferences.WikiLocation = getStandardDirectory("Wiki",baseDirectory); ;
                Settings.userPreferences.DiagramLocation = baseDirectory + @"png\";
                if (string.IsNullOrWhiteSpace(Settings.userPreferences.ZIB_RTFLocation)) Settings.userPreferences.ZIB_RTFLocation = getStandardDirectory("RTF", baseDirectory); ;
            }

            /// <summary>
            /// SetTargetDirectory zet de output folders zoals gedefinieerd in de configuratiefile om naar een op te geven basisfolder (baseDirectory).
            /// Dit wordt gebruikt om een set bestanden die voor een testomgeving aangemaakt wordt op te slaan voor hergebruik: opladen naar 
            /// een produktieomgeving, als het blijkt dat de bestanden goed zijn. Dit voorkomt dat de bestanden opnieuw gegenereerd moeten worden.
            /// </summary>
            /// <param name="baseDirectory">De 'root' van de folders waar de resultaten naar toe worden geschreven</param>
            /// <returns>String array met de paden van de huidige uitvoerfolders</returns>
            private string[] setTargetDirectories(string baseDirectory)
            {
                string[] currentDirs = new string[]
                {
                    Settings.userPreferences.VS_RTFLocation,
                    Settings.userPreferences.XMLLocation,
                    Settings.userPreferences.XLSLocation,
                    Settings.userPreferences.WikiLocation,
                    Settings.userPreferences.DiagramLocation,
                    Settings.userPreferences.ZIB_RTFLocation
                };

                Settings.userPreferences.XMLLocation = getStandardDirectory("XML", baseDirectory);
                Settings.userPreferences.XLSLocation = getStandardDirectory("XLS", baseDirectory);
                Settings.userPreferences.WikiLocation = getStandardDirectory("Wiki", baseDirectory);
                Settings.userPreferences.DiagramLocation = getStandardDirectory("PNG", baseDirectory);
                Settings.userPreferences.ZIB_RTFLocation = getStandardDirectory("RTF", baseDirectory);

                return currentDirs;
            }

            /// <summary>
            /// ResetTargetDirectories zet de output folders terug naar de eerder, met setTargetDirectories, opgeslagen waarden.
            /// </summary>
            /// <param name="previousDirs">String array met de eerder opgeslagen namen van outputfolders</param>
            private void resetTargetDirectories (string[] previousDirs)
            {
                Settings.userPreferences.VS_RTFLocation = previousDirs[0];
                Settings.userPreferences.XMLLocation = previousDirs[1];
                Settings.userPreferences.XLSLocation = previousDirs[2];
                Settings.userPreferences.WikiLocation = previousDirs[3];
                Settings.userPreferences.DiagramLocation = previousDirs[4];
                Settings.userPreferences.ZIB_RTFLocation = previousDirs[5];
            }


            public string getStandardDirectory(string outputType, string baseDir)
            {
                string stdDir = "";
                switch (outputType)
                {
                    case "XML":
                        stdDir = baseDir + @"xml\";
                        break;
                    case "XLS":
                        stdDir = baseDir + @"xls\";
                        break;
                    case "Wiki":
                        stdDir = baseDir + @"WikiPages\";
                        break;
                    case "RTF":
                        stdDir = baseDir + @"zrtf\";
                        break;
                    case "PNG":
                        stdDir = baseDir + @"png\";
                        break;
                    default:
                        stdDir = baseDir;
                        break;
                }
                return stdDir;

            }

            public void createSessionBaseDirectory()
            {
                if (!Directory.Exists(locationBase)) Directory.CreateDirectory(locationBase);
                Settings.userPreferences.sessionBase = locationBase + DateTime.Now.ToString("yyyyMMddHHmmss") + @"\";
                if (!Directory.Exists(Settings.userPreferences.sessionBase)) Directory.CreateDirectory(Settings.userPreferences.sessionBase);
            }

            public void createDirectories()
            {
                if (!Directory.Exists(Settings.userPreferences.VS_RTFLocation)) Directory.CreateDirectory(Settings.userPreferences.VS_RTFLocation);
                if (!Directory.Exists(Settings.userPreferences.WikiLocation)) Directory.CreateDirectory(Settings.userPreferences.WikiLocation);
                if (!Directory.Exists(Settings.userPreferences.XMLLocation)) Directory.CreateDirectory(Settings.userPreferences.XMLLocation);
                if (!Directory.Exists(Settings.userPreferences.XLSLocation)) Directory.CreateDirectory(Settings.userPreferences.XLSLocation);
                if (!Directory.Exists(Settings.userPreferences.DiagramLocation)) Directory.CreateDirectory(Settings.userPreferences.DiagramLocation);
                if (!Directory.Exists(Settings.userPreferences.ZIB_RTFLocation)) Directory.CreateDirectory(Settings.userPreferences.ZIB_RTFLocation);
            }

            private List<string> getRtfTemplates(EA.Repository r)
            {
                List<string> rtfTemplates = new List<string>();
                string strTemplates = "";
                try
                {
                    strTemplates = r.SQLQuery("SELECT DocName FROM t_document AS d WHERE d.ElementType = 'SSDOCSTYLE'");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Fout " + ex.ToString());
                }

                XDocument doc = XDocument.Parse(strTemplates);
                foreach (XElement t in doc.Descendants("DocName"))
                    rtfTemplates.Add(t.Value.ToString());
                return rtfTemplates;
            }

            private bool IsZibRepository()
            {
                string zibRepositoryFlag = r.Models.GetAt(0).Packages.GetAt(0).Element.TaggedValuesEx.GetByName("HCIM::ZIBRepository")?.Value ?? "False";
                return zibRepositoryFlag.ToTitleCase() == "True";
            }

            private void readReleaseTags()
            {
                if (releaseList.Count == 0) return;  // Configuratiefile fout die al gemeld is
                Settings.releasecontext.releaseYear = 0;
                Settings.releasecontext.preReleaseNumber = -1;
                Settings.releasecontext.releaseType = "Release";
                EventWithStringArgs e = null;

                foreach (EA.TaggedValue tag in r.Models.GetAt(0).Packages.GetAt(0).Element.TaggedValuesEx)
                    switch (tag.Name)
                    {
                        case "HCIM::ReleaseType":
                            Settings.releasecontext.releaseType = tag.Value; ;
                            break;
                        case "HCIM::ReleaseYear":
                            Settings.releasecontext.releaseYear = int.Parse(tag.Value); ;
                            break;
                        case "HCIM::PreReleaseNumber":
                            Settings.releasecontext.preReleaseNumber = int.Parse(tag.Value);
                            break;
                        default:
                            break;
                    }

               // 3-10-21 Test op aanwezigheid project tags en waarschuwing toegevoegd

                if (Settings.releasecontext.releaseYear == 0 || Settings.releasecontext.preReleaseNumber == -1)
                {
                    MessageBox.Show("De projectfile bevat geen informatie over publicatie jaar en/of pre-publicatienummer\r\n"
                               + "Maak deze eerst aan om de goede publicatie gegevens in de zib's te verwerken\r\n"
                               + "De configuratie is niet aangepast.", "Let op!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            
                // 3-10-21 Uit de if loop gehaald en naar voren verplaatst
                // 23-12-2025 test op prereleaseNumber toegevoegd
                var releaseData = releaseList.Where(x => x[0] == Settings.releasecontext.releaseYear.ToString() && x[4] == Settings.releasecontext.preReleaseNumber.ToString())?.FirstOrDefault();
                if ((Settings.releasecontext.releaseYear != 0 && Settings.releasecontext.releaseYear.ToString() != Settings.zibcontext.publicatie) ||
                    (Settings.releasecontext.preReleaseNumber !=Settings.zibcontext.PreReleaseNumber))
                {
                    MessageBox.Show("De publicatie in de gekozen configuratie komt niet\r\novereen met de publicatie in de projectfile\r\n"
                        + "Het publicatiejaar en overige publicatiegegevens worden aangepast aan de projectfile gegevens\r\n"
                        + "Projectfile meldt '" + Settings.releasecontext.releaseYear.ToString()
                        + (Settings.releasecontext.preReleaseNumber == 0 ? "" : ("-" + Settings.releasecontext.preReleaseNumber.ToString()))
                        + "' en de configuratie '" + Settings.zibcontext.publicatie
                        + (Settings.zibcontext.PreReleaseNumber == 0 ? "" : ("-" + Settings.zibcontext.PreReleaseNumber.ToString())) + "'"
                        + "\r\nPas eventueel de configuratie aan.", "Let op!", MessageBoxButtons.OK,
                        MessageBoxIcon.Warning); ;

                    if (releaseData == null)
                    {
                        MessageBox.Show("Er is geen publicatie configuratie gevonden die\r\novereenkomt met de publicatie in de projectfile\r\n"
                                   + "Maak deze eerst aan om de goede publicatie gegevens in de zib's te verwerken\r\n"
                                   + "De configuratie wordt op de laatste (pre-)publicatie gezet.", "Let op!", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        // Geen publicatie gevonden, default wordt de laatste. Negeer bovendien de projectfile gegevens

                        releaseData = releaseList[releaseList.Count -1] ;
                        Settings.releasecontext.releaseYear = int.Parse(releaseData[0]);
                        Settings.releasecontext.preReleaseNumber = int.Parse(releaseData[4]);
                        Settings.releasecontext.releaseType = releaseData[4] == "0" ? "Release" : "PreRelease";
                        //                        return;
                    }

                    Settings.zibcontext.publicatie = Settings.releasecontext.releaseYear.ToString();
                    Settings.zibcontext.ReleaseInfo = releaseData[1];
                    Settings.zibcontext.PreReleaseNumber =int.Parse(releaseData[4]);
                    Settings.wikicontext.ArtDecorRepository = releaseData[2];
                    Settings.wikicontext.ArtDecorProjectOID = releaseData[3];

                    e = new EventWithStringArgs(Settings.zibcontext.publicatie + (Settings.zibcontext.PreReleaseNumber == 0 ? "" : ("-" + Settings.zibcontext.PreReleaseNumber.ToString())));
                    Config_ReleaseChanged(this, e);
                    Settings.saveToFile(configName, locationBase);
                }
                
                // Initieel kunnen de configuraties en releaselist dat nog uiteen lopen. Corrigeer dat en sla de correcte data op
                if (Settings.zibcontext.PreReleaseNumber != int.Parse(releaseData[4]))
                {
                    Settings.zibcontext.PreReleaseNumber = int.Parse(releaseData[4]);
                    e = new EventWithStringArgs(Settings.zibcontext.publicatie + (Settings.zibcontext.PreReleaseNumber == 0 ? "" : ("-" + Settings.zibcontext.PreReleaseNumber.ToString())));
                    Config_ReleaseChanged(this, e);
                    Settings.saveToFile(configName, locationBase);
                }
                // Deze melding is overbodig geworden omdat deze test nu al in de eerste test zit
                /*
                if (Settings.zibcontext.PreReleaseNumber != Settings.releasecontext.preReleaseNumber)
                {
                    MessageBox.Show("De prerelease informatie van de configuratie en uit de projectfile komen niet overeen\r\n"
                               + "Projectfile meldt " + (Settings.releasecontext.preReleaseNumber == 0 ? "'Publicatie'" : ("'Pre-publicatie " + Settings.releasecontext.preReleaseNumber.ToString() + "'"))
                               + " en de configuratie " + (Settings.zibcontext.PreReleaseNumber == 0 ? "'Publicatie'" : ("'Pre-publicatie " + Settings.zibcontext.PreReleaseNumber.ToString() + "'"))
                               + "\r\nPas om de goede publicatie gegevens in de zib's te verwerken eerst de projectfile of de publicatie aan",
                                "Let op!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                */

                // 22-12-2025: Behouden prereleases: Settings.releasecontext.releaseFullName en Settings.application.MultipleReleasesFromYear om de wiki pagina naamgeving 
                // aan te passen zodat er meer dan een (pre-)publicatie per jaar mogelijk te maken en toch compatible met bestaande releases mogelijk te maken.

                Settings.releasecontext.releaseFullName = Settings.GetReleaseFullName(Settings.releasecontext.releaseYear, Settings.releasecontext.preReleaseNumber);

                return;
            }


            // ============================================
            // METHODEN AANGEROEPEN VANUIT DE MENU HANDLERS
            // ============================================

            public void OpenFile(string fileNaam)
            {
                string connectionString = "";
                string fileCopy = "";
                string fileLog = "";

                //                configName = Settings.readFromFile();
                this.Result.Clear();
                Settings.getConfigurationsList();
                frmConfigChoice confDiag = new frmConfigChoice();

                confDiag.ShowDialog();
                configName = confDiag.Config;

                if (configName == "") return;
                
                Settings.loadConfiguration(configName);
                this.lblConfigValue.Text = configName;
                createSessionBaseDirectory();
                setPreferences(Settings.userPreferences.sessionBase);
                createDirectories();

                //
                if (Path.GetExtension(fileNaam) == ".easvrlnk" || Path.GetExtension(fileNaam) == ".easvr")  // Bestand met een EA connectionstring voor een server database
                {
                    connectionString = File.ReadAllText(fileNaam);
                    if (connectionString.Contains("{0}"))  // we hebben een email account nodig voor inloggen op de database
                    {
                        connectionString = string.Format(connectionString, "@nictiz.nl");
                    }

                    // Maak een kopie door een server database transfer naar een eapx of qeax file database. Hiermee wordt verder gegaan
                    // Evt user management wordt meegekopieerd.
                    EA.Repository repository = new EA.Repository();
                    EA.Project project = repository.GetProjectInterface();
                    string tempRepositoryName = "ZIBS_temp";
                    fileCopy = Path.Combine(Settings.userPreferences.sessionBase, tempRepositoryName) + (bitness.Contains("32") ? ".eapx" : ".qeax");
                    fileLog = Path.Combine(Settings.userPreferences.sessionBase, tempRepositoryName) + ".log";
                    project.ProjectTransfer(connectionString, fileCopy, fileLog);
                    repository.Exit();
                }
                else
                {
                    connectionString = fileNaam;
                    // Maak een kopie voor evt. print van taal specifieke documenten
                    string file = Path.GetFileName(connectionString);
                    fileCopy = Path.Combine(Settings.userPreferences.sessionBase, file);
                    File.Copy(fileNaam, fileCopy);
                }

                StatusLabel.Text = "Opening Enterprise Architect repository ..........";
                EAStatus.Refresh();
                r = new EA.Repository();
                List<string> temp = new List<string>();
                StatusLabel.Text = "Reading Enterprise Architect file ..........";
                EAStatus.Refresh();
                Cursor.Current = Cursors.AppStarting;
                // 2-10-23: Als security enabled is moet een user en pwd opgegeven worden. Dit kan pas vastgesteld worden als de file open is: kip en ei.
                // Als patch het admin account gebruikt. Als security niet enabled is, gaat het ook goed.
                // Dit is natuurlijk wel foutgevoelig en eigenlijk niet gewenst. Gebruik van OpenID lijkt in de interop niet geregeld
                r.OpenFile2(fileCopy, "admin", "password");
                //r.OpenFile(fileCopy);
                // 26-05-23 Toegevoegd test op locks
                if (r.IsSecurityEnabled)
                {
                    if (HasLocks(r, out string errorMessage))
                    {
                        Result.Text += errorMessage + " \r\nVerwijder eerst de locks\r\nDe projectfile wordt nu gesloten";
                        sluitenToolStripMenuItem.PerformClick();
                        return;
                    }
                }
                // tot hier

                connectionString = r.ConnectionString;// Aangepast omdat DatabaseValid ref nu ref is. 
                if (!EAUtils.DatabaseValid(ref connectionString, out string dbType, out string dbConn))
                ///if (!EAUtils.DatabaseValid(r.ConnectionString, out string dbType, out string dbConn))
                    {
                        MessageBox.Show("Niet ondersteund type database: "+ dbConn + ": " + dbType, "Database fout",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    sluitenToolStripMenuItem.PerformClick();
                    return;
                }

                if (!IsZibRepository())
                {
                    MessageBox.Show("EA projectfile lijkt geen zib repository te zijn.", "Onbekend type project file",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    sluitenToolStripMenuItem.PerformClick();
                    return;
                }

                Cursor.Current = Cursors.Default;
                readReleaseTags();
                populateProjectView(saveState: false);
                StatusLabel.Text = "Ready";
                EAStatus.Refresh();

                //          voorkom dat een tweede file geopend wordt.
                this.openToolStripMenuItem.Enabled = false;

                //          File is geopend: maak de temp output directories aan

                var languageButtons = gBAction.Controls.OfType<ButtonBase>().Where(x => x.Name == "ccSingleLanguage");
                if (languageButtons.Count() > 0)
                    languageButtons.First().Text = zibActions["SingleLanguage"].description + " (" + Settings.zibcontext.pubLanguage.ToString() + ")";
                else
                    MessageBox.Show("Geen Language button gevonden"); 

                //          Lees de RFT Templates in
                templateList = getRtfTemplates(r);
                if (!templateList.Contains(Settings.zibcontext.zibTemplate))
                    Settings.zibcontext.zibTemplate = templateList[0];

                this.lblExampleDirValue.Text = Settings.userPreferences.ExampleLocation.EllipsisFilename(lblExampleDirValue.MaximumSize, lblExampleDirValue.Font);
                lblTemplateValue.Text = Settings.zibcontext.zibTemplate;
                lblReleaseLanguageValue.Text = Settings.zibcontext.pubLanguage.ToString();
                lblPrefixValue.Text = Settings.zibcontext.zibPrefix;

                lblReleaseValue.Text = Settings.zibcontext.publicatie + (Settings.zibcontext.PreReleaseNumber == 0 ? "" : ("-" + Settings.zibcontext.PreReleaseNumber.ToString()));
                lblWikiSiteValue.Text = Settings.wikicontext.wikiBaserurl;

                // geef de configuratie dialoog vrij
                this.configuratieToolStripMenuItem.Enabled = true;

                eapFileOpen = true;
                gBAction_CheckedChanged(btnActie, EventArgs.Empty);
            }

            // 26-05-23 test op locks toegevoegd
            private bool HasLocks(EA.Repository repository, out string errorMessage)
            {
                bool locks = true;
                int count;
                errorMessage = "";
                string queryReply;
                string SQLquery = "SELECT COUNT(*) AS Count FROM t_seclocks";
                try
                {
                    queryReply = repository.SQLQuery(SQLquery);
                }
                catch (Exception)
                {
                    //pository.EnsureOutputVisible("System");
                    //repository.WriteOutput("System", "ZIB EA Add-in - HasUserLocks: " + ex.ToString(), 0);
                    return true;
                }
                XDocument doc = XDocument.Parse(queryReply);
                bool succes = int.TryParse(doc.Descendants("Count")?.FirstOrDefault()?.Value ?? "-1", out count);
                if (succes)
                    if (count > 0 || count < 0)
                    {
                        locks = true;
                        if (count > 0)
                            errorMessage = "Er zijn locks gevonden";
                        else
                            errorMessage = "Door fouten is niet vast te stellen of er locks zijn";
                    }
                    else 
                        locks = false;
                else
                {
                    locks = true;
                    errorMessage = "Door fouten is niet vast te stellen of er locks zijn";
                }
                return locks;
            }




            private void clearElementWindows()
            {
                this.Result.Clear();
            }

            private void clearForm()
            {
                this.projectView.Nodes.Clear();
                //this.clearElementWindows();  Wat is het nut van deze method??
                this.lblFileNameValue.Text = "";
                this.lblSelectedValue.Text = "";
                this.lblZibLanguageValue.Text = "";
                this.lblTemplateValue.Text = "";
                this.lblReleaseValue.Text = "";
                this.lblConfigValue.Text = "";
                this.lblWikiSiteValue.Text = "";
                this.lblExampleDirValue.Text = "";
                this.lblReleaseLanguageValue.Text = "";
                this.lblPrefixValue.Text = "";
                //this.Result.Clear();

                var languageButtons = gBAction.Controls.OfType<Control>().Where(x => x.Name == "ccSingleLanguage");
                if (languageButtons.Count() > 0) languageButtons.First().Text = zibActions["SingleLanguage"].description;
            }


            public void deleteTempFiles()
            {
                if (!cbSaveTempDirs.Checked)
                {
                    if (Directory.Exists(Settings.userPreferences.sessionBase))
                        foreach (string file in Directory.GetFiles(Settings.userPreferences.sessionBase))
                            if (Path.GetExtension(file).ToLower() != ".log") File.Delete(file);
                    if (Directory.Exists(Settings.userPreferences.DiagramLocation)) Directory.Delete(Settings.userPreferences.DiagramLocation, true);
                    if (Directory.Exists(Settings.userPreferences.VS_RTFLocation)) Directory.Delete(Settings.userPreferences.VS_RTFLocation, true);
                }
            }
            public void deleteEmptyDirectories()
            {
                if (Directory.Exists(Settings.userPreferences.sessionBase))
                {
                    foreach (string subDir in Directory.GetDirectories(Settings.userPreferences.sessionBase))
                    {
                        bool wait = true;
                        while (wait)
                        {
                            try
                            {
                                if (Directory.Exists(subDir) && DirectoryIsEmpty(subDir)) Directory.Delete(subDir);
                                wait = false;                           }
                            catch (IOException e)
                            {
                                if ((uint)e.HResult == 0x80070020) //The process cannot access the file because it is being used by another process.
                                    wait = true;  // Eigenlijk nog een wait counter toepassen.
                                else
                                {
                                    wait = false;
                                    throw;
                                }
                            }
                        }
                    }
                    if (DirectoryIsEmpty(Settings.userPreferences.sessionBase)) Directory.Delete(Settings.userPreferences.sessionBase);
                }
            }

            private bool DirectoryIsEmpty( string dirPath)
            {
                if (Directory.GetFiles(dirPath).Length == 0 && Directory.GetDirectories(dirPath).Length == 0)
                    return true;
                else
                    return false;
            }

            public void emptyWorkDirectories()
            {
                if (Directory.Exists(Settings.userPreferences.sessionBase))
                {
                    try
                    {
                        foreach (string subDir in Directory.GetDirectories(Settings.userPreferences.sessionBase))
                        {
                            foreach (string file in Directory.GetFiles(subDir))
                                File.Delete(file);
                        }
                    }
                    catch 
                    { } //gaat het fout, jammer maar geen ramp
                }
            }


            /// <summary>
            /// Handmatig schalen van het configuratie dialog tbv andere resoluties dan 96 dpi omdat FOrms de schaling niet geweldig doet.
            /// Als er groupBoxes etc. bijkomen moeten deze ook hier suspend worden. Lastig, maar als dit deel in de designer staat wist MS het
            /// </summary>
            private void ScaleForm()
            {

                // gekopieerd uit Designer om layout tijdens de settings stil te leggen


                this.menuStrip1.SuspendLayout();
                this.contextMenuResult.SuspendLayout();
                this.EAStatus.SuspendLayout();
                ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
                this.splitContainer1.Panel1.SuspendLayout();
                this.splitContainer1.Panel2.SuspendLayout();
                this.splitContainer1.SuspendLayout();
                ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).BeginInit();
                this.splitContainer2.Panel1.SuspendLayout();
                this.splitContainer2.Panel2.SuspendLayout();
                this.splitContainer2.SuspendLayout();
                this.gbExecute.SuspendLayout();
                ((System.ComponentModel.ISupportInitialize)(this.pbLogo)).BeginInit();
                this.gbResult.SuspendLayout();
                this.gbOption.SuspendLayout();
                this.gbSettings.SuspendLayout();
                this.panel1.SuspendLayout();
                this.SuspendLayout();

                // DPI maakt de buttons veel te groot
                GroupBox[] groupBoxes = null;
                if (this.AutoScaleMode == AutoScaleMode.Dpi)
                {
                    groupBoxes = new GroupBox[] { gbExecute, gbResult };

                    foreach (GroupBox groupBox in groupBoxes)
                    {
                        // foreach button mag niet met ref
                        Button[] buttons = groupBox.Controls.OfType<Button>().ToArray();
                        for (int i = 0; i < buttons.Count(); i++)
                            ScaleButton(ref buttons[i]);
                    }
                }
                btnActie.Width = btCancel.Width;
                int startXButton = (gbExecute.Width - 2 * btnActie.Width - Scale(6)) / 2;
                btnActie.Location = new Point(startXButton, btCancel.Top);
                btCancel.Location = new Point(btnActie.Right + Scale(6), btCancel.Top);

                gbSettings.Location = new Point(gbSettings.Location.X, gbExecute.Location.Y + gbExecute.Size.Height + 3);
                if (gbOption.Size.Height > gBAction.Size.Height) gbOption.Size = new Size(gbOption.Size.Width, gBAction.Size.Height);

                int newSplitterDistance = splitContainer2.Height - (gbSettings.Location.Y + gbSettings.Size.Height + Scale(15));
                splitContainer2.SplitterDistance = newSplitterDistance;
                gbResult.Location = new Point(gbExecute.Location.X + gbExecute.Width + Scale(6), gbResult.Location.Y);
                int endOptionBox = gbOption.Location.X + gbOption.Size.Width;
                int endResultBox = gbResult.Location.X + gbResult.Size.Width;
                int endSettingsBox = gbSettings.Location.X + gbSettings.Size.Width;
                int shortestBox = Math.Min(endOptionBox, Math.Min(endResultBox, endSettingsBox));
                int largestBox = Math.Max(endOptionBox, Math.Max(endResultBox, endSettingsBox));
                if (shortestBox < gbResult.Location.X + btResultLocation.Location.X + btResultLocation.Width)
                {
                    shortestBox = gbResult.Location.X + btResultLocation.Location.X + btResultLocation.Width;
                    // voer de nieuwe size door en kijk of winforms deze aanpast.
                    gbResult.Size = new Size(shortestBox - gbResult.Location.X, gbResult.Size.Height);
                    shortestBox = gbResult.Location.X + gbResult.Size.Width;
                }
                gbOption.Size = new Size(shortestBox - gbOption.Location.X, gbOption.Size.Height);
                gbResult.Size = new Size(shortestBox - gbResult.Location.X, gbResult.Size.Height);
                gbSettings.Size = new Size(shortestBox - gbSettings.Location.X, gbSettings.Size.Height);
                this.Width = splitContainer1.Panel1.Width + splitContainer1.SplitterWidth + shortestBox + Scale(40);

                // gekopieerd uit Designer om layout tijdens de settings stil te leggen

                this.menuStrip1.ResumeLayout(false);
                this.menuStrip1.PerformLayout();
                this.contextMenuResult.ResumeLayout(false);
                this.EAStatus.ResumeLayout(false);
                this.EAStatus.PerformLayout();
                this.splitContainer1.Panel1.ResumeLayout(false);
                this.splitContainer1.Panel2.ResumeLayout(false);
                ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
                this.splitContainer1.ResumeLayout(false);
                this.splitContainer2.Panel1.ResumeLayout(false);
                this.splitContainer2.Panel1.PerformLayout();
                this.splitContainer2.Panel2.ResumeLayout(false);
                this.splitContainer2.Panel2.PerformLayout();
                ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).EndInit();
                this.splitContainer2.ResumeLayout(false);
                this.gbExecute.ResumeLayout(false);
                this.gbExecute.PerformLayout();
                ((System.ComponentModel.ISupportInitialize)(this.pbLogo)).EndInit();
                this.gbResult.ResumeLayout(false);
                this.gbResult.PerformLayout();
                this.gbOption.ResumeLayout(false);
                this.gbOption.PerformLayout();
                this.gbSettings.ResumeLayout(false);
                this.gbSettings.PerformLayout();
                this.panel1.ResumeLayout(false);
                this.panel1.PerformLayout();
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





            private void about(object o, EventArgs ea)
            {
                AboutBox1 aboutWindow = new AboutBox1();
                aboutWindow.StartPosition = FormStartPosition.CenterParent;
                aboutWindow.Show();
            }

            /// <summary>
            /// HULP CLASSES EN ENUMS
            /// </summary>
            /// 


            // ======================
            // Class actionData
            // ======================
            public class actionData
            {
                public string ctrlName { get; set; }
                public string description { get; set; }
                public zibAction function { get; set; }
                public bool isChecked { get; set; }
                public bool isBatchable { get; set; } 
                
                public actionData(string _ctrlName, string _description, zibAction _function, bool _isChecked, bool _isBatchable)
                {
                    ctrlName = _ctrlName;
                    description = _description;
                    function = _function;
                    isChecked = _isChecked;
                    isBatchable = _isBatchable;
                }      
            }


            // ======================
            // Enum PackageType
            // ======================

            public enum PackageType
            {
                other,
                zibcontainer,
                singlezib
            }
 
            private PackageType packageType(EA.Package p)
            {
                if (p.StereotypeEx == "DCM")
                    return PackageType.singlezib;
                else if (p.Packages.Count > 0 && p.Packages.GetAt(0).StereotypeEx == "DCM")
                    return PackageType.zibcontainer;
                else
                    return PackageType.other;
            }
        
        }
    }
}





