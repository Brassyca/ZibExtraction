using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.ComponentModel;
using Zibs.Configuration;
using Zibs.tekstManager;
using Zibs.ExtensionClasses;
using JiraJsonClient;

namespace Zibs.ZibExtraction
{
    public partial class ZibExtraction : Form
    {
        private ErrorType logLevel;

        /// <summary>
        /// Background worker event methodes
        ///
        /// </summary>

        private void zibExtractor_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            performExtraction(((zibExtractorPars)(e.Argument)).actionList, ((zibExtractorPars)(e.Argument)).options, ((zibExtractorPars)(e.Argument)).copyPath, ((zibExtractorPars)(e.Argument)).logLevel, worker, e);
        }


        private void performExtraction(List<zibAction> toDoList, Dictionary<string, bool> options, string copyPath, ErrorType _logLevel, BackgroundWorker worker, DoWorkEventArgs e)
        {
            logLevel = _logLevel;
            bool succes = false;
            int totalSteps = toDoList.Count * zibList.Count;
            int zibCount = 0;
            int stepCount = 0;
            bool inReleaseChecked = false;
            textLanguage checkLanguage;
            errorCount = 0;

            // Start de logging

            messageLog.AppendLine("=== Message log ZibExtraction sessie dd: " + string.Format("{0:d}, {0:t}", DateTime.Now) + "===\r\n");
            errorLog.AppendLine("=== Error log ZibExtraction sessie dd: " + string.Format("{0:d}, {0:t}", DateTime.Now) + "===\r\n");


            // Maak een backup van de Groep en publicatie config file bij zibreg of refreg acties en de optie backup aan staat.
            if (options["cbBackup"])
            {
                string fileName = Path.Combine(Settings.application.CommonConfigLocation, Settings.zibRegistryFileName);
                string bckFileName = Path.GetFileNameWithoutExtension(fileName) + "_Backup_" + DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetExtension(fileName);
                File.Copy(fileName, Path.Combine(Path.GetDirectoryName(fileName), bckFileName));
            }

            reportProgressText(ErrorType.general, "\r\n== Prefix check\r\n", worker);
            foreach (EA.Package _zib in zibList)
            {
                if (zibName.Prefix(_zib.Name) != Settings.zibcontext.zibPrefix)
                    reportProgressText(ErrorType.warning, "Zib " + _zib.Name + " heeft een afwijkende prefix: '"+ zibName.Prefix(_zib.Name) + "' i.p.v. '" +  Settings.zibcontext.zibPrefix + " en wordt daarom overgeslagen\r\n", worker);
            }
            reportProgressText(ErrorType.general, "== Einde prefix check\r\n", worker);


            // de feitelijke acties
            foreach (zibAction _action in toDoList)
            {

                // check eerst éénmalig of de zib + versie in deel van de publicatie is, behalve bij een zib registratie
                if (!inReleaseChecked)
                {
                    if (_action == registerZIB || _action == createWikiTOCpage || _action == uploadWikiFiles)
                        checkLanguage = textLanguage.Multi;
                    else if (_action == registerRefs || _action == createSingleLanguageZIB || _action == getZIB_XMLfiles)
                        checkLanguage = textLanguage.NL;
                    else
                        checkLanguage = Settings.zibcontext.pubLanguage;

                    if (checkLanguage != textLanguage.Multi)
                    {
                        bool errorsFound = false;
                        reportProgressText(ErrorType.general, "\r\n== Check of zibs tot de publicatie behoren\r\n", worker);
                        foreach (EA.Package _zib in zibList)
                        {
                            bool inRelease = IsInRelease(_zib, Settings.zibcontext.publicatie, checkLanguage);
                            if (!inRelease)
                            {
                                reportProgressText(ErrorType.warning, "Zib " + _zib.Name + " is niet geregistreerd als onderdeel van de " + checkLanguage.ToString() + " publicatie " + Settings.zibcontext.publicatie + " -----\r\n", worker);
                                errorsFound = true;
                            }
                        }
                        if (!errorsFound) reportProgressText(ErrorType.information, "Alle geselecteerde zibs zijn onderdeel van publicatie " + Settings.zibcontext.publicatie + "\r\n", worker);
                        reportProgressText(ErrorType.general, "== Einde check\r\n", worker);

                        inReleaseChecked = true;
                    }                            
                }

                stepCount++;
                foreach (EA.Package _zib in zibList)
                {
                    if (worker.CancellationPending)
                    {
                        e.Cancel = true;
                        return;
                    }
                    zibCount++;
                    reportBarLabel("Actie " + stepCount.ToString() + "/" + toDoList.Count + " op zib " + zibCount.ToString() + "/" + zibList.Count, worker);

                    if (zibName.Prefix(_zib.Name) != Settings.zibcontext.zibPrefix)
                    {
                        reportProgressBar((int)((float)(zibCount + (stepCount - 1) * zibList.Count) / (float)totalSteps * 100), worker);
                        continue;
                    }

                    succes = getZIB(_zib, _action, options, worker);
                    if (!succes) break;
                    reportProgressBar((int)((float)(zibCount + (stepCount - 1) * zibList.Count) / (float)totalSteps * 100), worker);
                }
                if (!succes) break;
                if (_action == createSingleLanguageZIB)
                {
                    rePopulateProjectView(worker);  //repaint de ProjectView maar één maal aan het einde van de actie om bevriezing van de UI te voorkomen
                }
                zibCount = 0;
                if(_action == registerZIB || _action == registerRefs)
                {
                    Settings.XGenConfig.Save(true);
                    // Lees de zib registry opnieuw in
                    Settings.XGenConfig.Reload();           //26-11-2021
                }
            }
            // Kopieer als de save checkbox gechecked is de resultaten naar de opgegeven locatie.
            //               if (cbSave.Checked) copyFiles(tbResultLocation.Text + @"\");
            if (options["cbSave"]) copyFiles(tbResultLocation.Text + @"\", worker);

            // Schrijf de logging naar de session temp directory

            string fullLogging = Path.Combine(Settings.userPreferences.sessionBase, "Fulllog_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".log");
            if (!File.Exists(fullLogging))
            {
                messageLog.AppendLine("\r\n=== Einde message log ZibExtraction sessie dd: " + string.Format("{0:d}, {0:t}", DateTime.Now) + "===\r\n");
                File.WriteAllText(fullLogging, messageLog.ToString(), Encoding.UTF8);
            }
            string errorLogging = Path.Combine(Settings.userPreferences.sessionBase, "Errorlog_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".log");
            if (!File.Exists(errorLogging))
            {
                errorLog.AppendLine("\r\n=== Einde error log ZibExtraction sessie dd: " + string.Format("{0:d}, {0:t}", DateTime.Now) + "===\r\n");
                File.WriteAllText(errorLogging, errorLog.ToString(), Encoding.UTF8);
            }
            
            if (errorCount > 0)
            {
                reportProgressText(ErrorType.general, "\r\nAantal 'warnings' of 'errors' is: " + errorCount.ToString() + "\r\n", worker);
            }
            else
            {
                reportProgressText(ErrorType.general, "\r\nGeen 'warnings' of 'errors' gevonden\r\n", worker);
            }

            MessageBox.Show("De gekozen acties zijn uitgevoerd", "Voltooid", MessageBoxButtons.OK, MessageBoxIcon.Information);
            reportProgressBar(0, worker);
            reportBarLabel("", worker);
        }

        private void zibExtractor_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.UserState == null)
            {
                this.progressBar1.Value = e.ProgressPercentage;
            }
            else
            {
                if (!string.IsNullOrEmpty(((reportPars)e.UserState).state))
                {
                    this.StatusLabel.Text = ((reportPars)e.UserState).state;
                    EAStatus.Refresh();
                }

                if (((reportPars)e.UserState).rePopulate)
                {
                    populateProjectView(saveState: true);
                }

                if (!string.IsNullOrEmpty(((reportPars)e.UserState).state))
                {
                    this.StatusLabel.Text = ((reportPars)e.UserState).state;
                    EAStatus.Refresh();
                }

                if (!string.IsNullOrEmpty(((reportPars)e.UserState).errorMessage))
                {
                    MessageBox.Show(((reportPars)e.UserState).errorMessage, "Foutmelding", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                if (!string.IsNullOrEmpty(((reportPars)e.UserState).informationMessage.Message))
                {
                    if (((reportPars)e.UserState).addLine)
                    {
                        ErrorType _type = ((reportPars)e.UserState).informationMessage.Type;
                        string _message = (_type == ErrorType.general ? "" : _type.ToString().ToTitleCase() + ": ") + ((reportPars)e.UserState).informationMessage.Message;
                        if (_type <= logLevel || _type == ErrorType.general) Result.AppendText(_message);
                        messageLog.Append(_message);
                        if (_type == ErrorType.disaster || _type == ErrorType.error || _type == ErrorType.warning)
                        {
                            errorLog.Append(_message);
                            errorCount++;
                        }
                    }
                    else
                    {
                        clearElementWindows();
                        string[] stringSeparators = new string[] { "\r\n" };
                        this.Result.Lines = ((reportPars)e.UserState).informationMessage.Message.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);
                    }
                }

                if (!string.IsNullOrEmpty(((reportPars)e.UserState).labelMessage))
                {
                    this.lblProgressBar.Text = (((reportPars)e.UserState).labelMessage);
                }
            }
        }

        private void zibExtractor_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // First, handle the case where an exception was thrown.
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message, "Foutmelding", MessageBoxButtons.OK, MessageBoxIcon.Error );
                this.StatusLabel.Text = "Error";
                EAStatus.Refresh();
            }
            else if (e.Cancelled)
            {
                // Next, handle the case where the user canceled 
                // the operation.
                // Note that due to a race condition in 
                // the DoWork event handler, the Cancelled
                // flag may not have been set, even though
                // CancelAsync was called.

                this.StatusLabel.Text = "Canceled";
                EAStatus.Refresh();
                Result.AppendText("\r\nActies door gebruiker geannuleerd\r\n");
                this.lblProgressBar.Text = "";
                DialogResult answ = MessageBox.Show("Resultaten tot nu toe bewaren?", "Acties geannuleerd", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (answ == DialogResult.No)
                    emptyWorkDirectories();
            }
            else
            {
                // Finally, handle the case where the operation 
                // succeeded.
                this.StatusLabel.Text = "Ready";
                EAStatus.Refresh();
            }

            this.progressBar1.Value = 0;
            this.lblProgressBar.Text = "";

            // Disable the Cancel button.
            this.btCancel.Enabled = false;

            // Enable the Check button.
            this.btnActie.Enabled = true;

        }


        /// <summary>
        /// ACTIONS
        /// worden aangeroepen via de backgroundworker 
        /// </summary>

        public bool getZIB(EA.Package p, zibAction ZibAction, Dictionary<string, bool> options, object sender)

        {
            bool succes = false;
            zib = new ZIB(p, r, getSummaryByIssue);
            zib.NewMessage += Zib_NewMessage;
            reportStatus("Processing  " + zib.Fullname + " ..........", sender);
            zibPublishLangauge = zib.getZibPublishLanguage();
            succes = ZibAction(options, sender);
            reportStatus("Ready", sender);
            return succes;
        }

        // =================
        // Action: text/wiki
        // =================

        private bool getZIB_sections(Dictionary<string, bool> options, object sender)
        {
            reportProgressText(ErrorType.general, "\r\n== Aanmaken wiki/text bestanden van " + zib.Name + "\r\n", sender);
            bool languageOk = actionLanguageCheck("wiki", sender);
            if (!languageOk)
            {
                reportProgressText(ErrorType.error, "== Aanmaken wiki/text bestanden gestopt\r\n", sender);
                return false;
            }

            // Maak eerst de pagina met issues aan als die er niet is
            if (issueDictionary.Count() == 0)
                getBitsIssues(sender);

            outputType o = outputType.wiki;
            List<string> temp = new List<string>();
            zib.getSections(ref temp, o, this.zibPublishLangauge);
            if (options["cBwikiPreview"])
            {
                StringBuilder builder = new StringBuilder();
                foreach (string line in temp) // Loop through all strings
                {
                    builder.Append(line);
                }
                reportAll(builder.ToString(), sender);
                reportProgressText(ErrorType.general, "\r\n", sender);
            }
            reportProgressText(ErrorType.information, "Aangemaakt: wiki/text bestand van " + zib.Name + "\r\n", sender);
            reportProgressText(ErrorType.general, "== Einde aanmaken wiki/text bestanden\r\n", sender);
            return true;
        }

        // ========================================
        // text/wiki : ophalen issue info uit Bits
        // ========================================

        private bool getBitsIssues(object sender)
        {
            reportProgressText(ErrorType.information, "Issue gegevensophalen uit Bits\r\n", sender);
            /*            jira Bits = new jira { BaseUrl = new Uri(Settings.bitscontext.bitsBaserurl) };
                        Bits.NewMessage += Bits_NewMessage;
                        Bits.bitsIssues(ref issueDictionary); */

            JiraClient Bits = new JiraClient
            {
                BaseUrl = new Uri(Settings.bitscontext.bitsBaserurl),
                User = Settings.bitscontext.bitsUser,
                Passwd = Settings.bitscontext.bitsPassword,
                ClientName = "Nictiz bot client",
                MaxIssuesPerPage = Settings.bitscontext.maxIssuesPerPage
            };

            Bits.NewMessage += Bits_NewMessage;
            Bits.bitsIssues(Settings.bitscontext.bitsStatus.Split(';').Select(x=>x.Trim()).ToArray(), ref issueDictionary);

            return true;
        }

/*        private void Bits_NewMessage(object sender, EventWithStringArgs e)
        {
            reportProgressText(e.Text + "\r\n", zibExtractor);  ////niet netjes
        }
*/
        private void Bits_NewMessage(object sender, ErrorMessageEventArgs e)
        {
            reportProgressText(e.MessageType, e.MessageText + "\r\n", zibExtractor);  ////niet netjes
        }

        public string getSummaryByIssue(string issue)
        {
            string summary = "";
            bool succes = false;
            if (issueDictionary.Count > 0)
                succes = issueDictionary.TryGetValue(issue, out summary);
            else
                return "Issue summaries niet beschikbaar";
            if (succes)
                return summary;
            else
                return "Geen informatie";
        }


        // =================
        // Action: XML
        // =================

        public bool getZIB_XMLfiles(Dictionary<string, bool> options, object sender)
        {
            outputType xmlType;
            reportProgressText(ErrorType.general, "\r\n== Aanmaken XML bestanden van " + zib.Name + "\r\n", sender);
            bool languageOk = actionLanguageCheck("xmi", sender);
            if (!languageOk)
            {
                reportProgressText(ErrorType.error, "== Aanmaken XML bestanden gestopt\r\n", sender);
                return false;
            }

            EA.Package informationModel;
            if (options["cBArtDecor"])
                xmlType = outputType.xmlAD;
            else
                xmlType = outputType.xml;

            informationModel = zib.EApackage.Packages.GetByName(sectionType.Information_Model.ToString().Replace('_', ' '));
            zib.getValueSets(informationModel, xmlType);
            reportProgressText(ErrorType.information, "Aangemaakt: XML bestand van de waardelijsten van " + zib.Name + "\r\n", sender);
            zib.getXMI(this.zibPublishLangauge);
            reportProgressText(ErrorType.information, "Aangemaakt: XMI bestand van " + zib.Name + "\r\n", sender);
            reportProgressText(ErrorType.general, "== Einde maken XML bestanden\r\n", sender);
            return true;
        }

        // ===================
        // Action: upload wiki
        // ===================

        public bool uploadWikiFiles(Dictionary<string, bool> options, object sender)
        {
            string[] currentDirs = null; ;
            if (uploadDone) return true;
            reportProgressText(ErrorType.general, "\r\n== Upload Wiki bestanden\r\n", sender);

            if (options["cbReuse"])
            {
                currentDirs = setTargetDirectories(tbResultLocation.Text + @"\");
                reportProgressText(ErrorType.general, "Aanpassen bestandslocatie naar " + tbResultLocation.Text + "\r\n", sender);
            }
            string adres = Settings.wikicontext.wikiBaserurl;
            string apiurl = "api.php";
            Wiki zibWiki = new Wiki { BaseUrl = new Uri(adres), RequestUrl = apiurl };
            zibWiki.NewMessage += zibWiki_NewMessage;

            string pagesToPurge;
            pagesToPurge = "";
            if (Settings.doPurgePages)
            {
                List<string[]> zibsInRelease = Settings.XGenConfig.getZibsInRelease(Settings.zibcontext.publicatie, Settings.zibcontext.pubLanguage);
                if (zibsInRelease.Count > 0)
                {
                    StringBuilder sb = new StringBuilder();
                    foreach (string[] zibEntry in zibsInRelease)
                    {
                        sb.Append(zibName.wikiLink(zibEntry[1] + "-v" + zibEntry[2], Settings.zibcontext.pubLanguage) + "|");
                    }
                    sb.Length = sb.Length - 1;
                    pagesToPurge = sb.ToString();
                }
            }
            zibWiki.uploadWikiAll(pagesToPurge);
            uploadDone = true;
            if (options["cbReuse"])
            {
                resetTargetDirectories(currentDirs);
                reportProgressText(ErrorType.general, "Terugzetten bestandlocatie naar default locatie\r\n", sender);
            }

            reportProgressText(ErrorType.general, "== Einde upload Wiki bestanden\r\n", sender);
            return true;
        }
        private void zibWiki_NewMessage(object sender, ErrorMessageEventArgs e)
        {
            reportProgressText(e.MessageType, e.MessageText + "\r\n", zibExtractor);
        }



        // =================
        // Action: DOCX
        // =================

        public bool getZIB_RTFfile(Dictionary<string, bool> options, object sender)
        {
            reportProgressText(ErrorType.general, "\r\n== Aanmaken " + (Settings.PDFFormat ? "PDF" : "DOCX") + " bestanden van " + zib.Name + "\r\n", sender);
            bool languageOk = actionLanguageCheck("zrtf", sender);
            if (!languageOk)
            {
                reportProgressText(ErrorType.error, "== Aanmaken " + (Settings.PDFFormat ? "PDF" : "DOCX") + " bestanden gestopt\r\n", sender);
                return false;
            }
            zib.printToRTF(this.zibPublishLangauge);
            reportProgressText(ErrorType.information, "Aangemaakt: " + (Settings.PDFFormat ? "PDF" : "DOCX") + " bestand van " + zib.Name + "\r\n", sender);
            reportProgressText(ErrorType.general, "== Einde aanmaken " + (Settings.PDFFormat ? "PDF" : "DOCX") + " bestanden\r\n", sender);
            return true;
        }

        // =================
        // Action: XLS
        // =================

        public bool createZIB_XLSfile(Dictionary<string, bool> options, object sender)
        {
            reportProgressText(ErrorType.general, "\r\n== Aanmaken spreadsheets van "+ zib.Name + "\r\n", sender);
            bool languageOk = actionLanguageCheck("xls", sender);
            if (!languageOk)
            {
                reportProgressText(ErrorType.error, "== Aanmaken spreadsheets gestopt\r\n", sender);
                return false;
            }
            List<string> temp = new List<string>();
            zib.getSections(ref temp, outputType.xls, this.zibPublishLangauge);
            if (options["cBwikiPreview"])
            {
                StringBuilder builder = new StringBuilder();
                foreach (string line in temp) // Loop through all strings
                {
                    builder.Append(line);
                }
                reportAll(builder.ToString(), sender);
                reportProgressText(ErrorType.general, "\r\n", sender);
            }
            reportProgressText(ErrorType.information, "Aangemaakt: XLS bestand van " + zib.Name + "\r\n", sender);
            reportProgressText(ErrorType.general, "== Einde aanmaken spreadsheets\r\n", sender);
            return true;
        }

        // =======================
        // Action: single language
        // =======================

        public bool createSingleLanguageZIB(Dictionary<string, bool> options, object sender)
        {
            bool succes = true;
            reportProgressText(ErrorType.general, "\r\n== Aanmaken ééntalige versie van " + zib.Name + "\r\n", sender);
            if (zib.Language == textLanguage.Multi)
            {
                if (zib.getSingleLanguageVersion())
                {
                    reportProgressText(ErrorType.information, "Aangemaakt: Eéntalige " + Settings.zibcontext.pubLanguage + " versie van " + zib.Fullname + "\r\n", sender);
                    this.zibPublishLangauge = Settings.zibcontext.pubLanguage;
                    if (zib.resizeDiagramTextboxes())
                        reportProgressText(ErrorType.information, "Hoogte van de Constraint en Notes tekstboxen van " + zib.Fullname + " aangepast\r\n", sender);
                    else
                    {
                        reportProgressText(ErrorType.warning, "Aanpassen hoogte van de tekstboxen van " + zib.Fullname + " mislukt!\r\n", sender);
                        succes = false;
                    }
                }
                else
                {
                    reportProgressText(ErrorType.error, "Aangemaak ééntalige " + Settings.zibcontext.pubLanguage + " versie van " + zib.Fullname + " mislukt!\r\n", sender);
                    succes = false;
                }

            }
            else
                reportProgressText(ErrorType.warning, zib.Fullname + " is al een ééntalige " + this.zibPublishLangauge + " versie\r\n", sender);
            reportProgressText(ErrorType.general, "== Einde aanmaken ééntalige versie\r\n", sender);
            //           populateProjectView(saveState: true);
            // rePopulateProjectView(sender); verplaatst naar boven aan het einde van de actie loop
            return succes;
        }


        // =======================
        // Action: registerZib
        // =======================

        public bool registerZIB(Dictionary<string, bool> options, object sender)
        {
            int succes = 0;
            bool doContinue = false;
            reportProgressText(ErrorType.general, "\r\n== Registeren Zib's\r\n", sender);

            doContinue = actionLanguageCheck("register", sender);
            if (!doContinue)
            {
                reportProgressText(ErrorType.error, "== Registreren zibs gestopt\r\n", sender);
                return false;
            }

            doContinue = true;
            succes = zib.registerZib(Settings.forceZibRegistration);
            if ((succes & 255) == 3)
                reportProgressText(ErrorType.information, zib.Name + " opnieuw geregistreerd\r\n", sender);
            else if ((succes & 255) == 2)
                reportProgressText(ErrorType.information, zib.Name + " is al geregistreerd\r\n", sender);
            else if ((succes & 255) == 1)
                reportProgressText(ErrorType.information, zib.Name + " succesvol geregistreerd\r\n", sender);
            else if ((succes & 255) == 0)
            {
                reportProgressText(ErrorType.error, zib.Name + ": registratie mislukt!\r\n", sender);
                doContinue = false;
            }
            else if ((succes & 255) == 99)
            {
                reportProgressText(ErrorType.error, zib.Name + ": niet geregistreerd, fout in Id of naam. Test zib!\r\n", sender);
                doContinue = false;
            }
            else
                reportProgressText(ErrorType.error, "Onbekende resultcode" + (succes & 255).ToString(), sender);


            if (succes == 1 || succes == 3 ) Settings.saveRegistryFile(false);
            reportProgressText(ErrorType.general, "== Einde registeren Zib's\r\n", sender);

            return doContinue;
        }

        // =======================
        // Action: registerRefs
        // =======================

        public bool registerRefs(Dictionary<string, bool> options, object sender)
        {
            int succes = 0;
            bool doContinue = false;
            reportProgressText(ErrorType.general, "\r\n== Registeren verwijzingen\r\n", sender);

            doContinue = actionLanguageCheck("references", sender);
            if (!doContinue)
            {
                reportProgressText(ErrorType.error, "== Registreren verwijzigingen gestopt\r\n", sender);
                return false;
            }

            doContinue = true;
            succes = zib.registerReferences(Settings.forceZibRegistration);
            if (succes == 11)
                reportProgressText(ErrorType.error, "Release niet gevonden. Maak deze eerst aan\r\n", sender);
            else if (succes == 12)
                reportProgressText(ErrorType.error, "Zib niet in release gevonden. Voeg deze eerst aan de release toe\r\n", sender);
            else if (succes == 13)
                reportProgressText(ErrorType.error, "Eén of meer zib's waar naar verwezen wordt, zijn geen deel van de publicatie " + Settings.zibcontext.publicatie + "\r\n", sender);
            else if (succes == 4)
                reportProgressText(ErrorType.information, "De bouwsteen heeft geen verwijzingen\r\n", sender);
            else if (succes == 3)
                reportProgressText(ErrorType.information, "Verwijzingen opnieuw geregistreerd\r\n", sender);
            else if (succes == 2)
                reportProgressText(ErrorType.information, "Verwijzingen zijn al geregistreerd\r\n", sender);
            else if (succes == 1)
                reportProgressText(ErrorType.information, "Verwijzingen succesvol geregistreerd\r\n", sender);
            else if (succes == 0)
            {
                reportProgressText(ErrorType.error, "Registratie verwijzingen mislukt!\r\n", sender);
                doContinue = false;
            }
            else
                reportProgressText(ErrorType.error, "Onbekende resultcode" + succes.ToString(), sender);

            if (succes == 1 || succes == 3) Settings.saveRegistryFile(false);

            reportProgressText(ErrorType.general, "== Einde registeren verwijzingen\r\n", sender);
            return doContinue;
        }


        // ======================================
        // Action: Aanmaken TOC en release pages
        // ======================================

        private bool createWikiTOCpage(Dictionary<string, bool> options, object sender)
        {
            if (tocDone) return true;
            reportProgressText(ErrorType.general, "\r\n== Aanmaken inhoudsopgave publicatie " + Settings.zibcontext.publicatie + " en issue pagina\r\n", sender);

            string transcludeName;

            // Maak ook de pagina met issues aan
            if (issueDictionary.Count() == 0)
                getBitsIssues(sender);

            textManager tm = new textManager("ZibExtractionLabels.cfg");
            if (tm.dictionaryOK == false)
            {
                reportProgressText(ErrorType.error, "getSections: Geen labeldictionary beschikbaar", sender);
                reportProgressText(ErrorType.error, "== Aanmaken inhoudsopgave gestopt\r\n", sender);
                tocDone = true;
                return false;
            }
            tm.Language = Settings.zibcontext.pubLanguage;


            StringBuilder sb = new StringBuilder();
            int rowCount;
            sb.AppendLine("== " + (Settings.releasecontext.releaseType == "Release" ? tm.getWikiLabel("rlHeader") : tm.getWikiLabel("rlPreHeader")) + " " + Settings.zibcontext.publicatie + (Settings.releasecontext.preReleaseNumber >0 ? ("-" + Settings.releasecontext.preReleaseNumber.ToString()): "") + " ==");
            List<string[]> zibsInRelease;
            zibsInRelease = Settings.XGenConfig.zibInReleasePerGroup(Settings.zibcontext.publicatie, Settings.zibcontext.pubLanguage);
            if (zibsInRelease.Count == 0) return true;

            int colCount = 4;
            var groups = zibsInRelease.Select(z => z[0]).ToList().Distinct().OrderBy(x => x);
            foreach (string group in groups)
            {
                List<string[]> groupList = zibsInRelease.Where(z => z[0] == group).OrderBy(x => x[3]).ToList();
                rowCount = (int)Math.Ceiling((double)groupList.Count / colCount);
                sb.AppendLine("==== " + tm.getWikiLabel("rlGroup") + ": " + group + ", " + tm.getWikiLabel("rlNumber") + ": " + groupList.Count + " ====");
                sb.AppendLine("{| width=\"1000\" style=\"background-color:#F8F8FF\"");
                int iRow = 0;
                while (iRow < rowCount)
                {
                    sb.AppendLine("|-");
                    for (int iCol = 0; iCol < colCount; iCol++)
                    {
                        int index = iRow + iCol * rowCount;
                        if (index < groupList.Count)
                        {
                            string zibname = groupList[index][3] + "-v" + groupList[index][2];
                            sb.AppendLine(((iRow == 0) ? "|width=\"25%\" |" : "|") + "[[" + zibName.wikiLink(zibname, Settings.zibcontext.pubLanguage) + "|" + zibname + "]]");
                        }
                        else
                            sb.AppendLine("|");
                    }
                    iRow++;
                }
                sb.AppendLine("|}");
            }
            string fileName = tm.getLabel("wikiReleasePage") + "_" + Settings.zibcontext.publicatie + "(" + Settings.zibcontext.pubLanguage.ToString() + ")_Section" + Settings.wikicontext.tocSection + ".wiki";
            File.WriteAllText(Path.Combine(Settings.userPreferences.WikiLocation, fileName), sb.ToString());

            // Schrijf de transclude pagina met release informatie.
            transcludeName = "HCIMReleases(" + Settings.zibcontext.pubLanguage + ").template";
            File.WriteAllText(Path.Combine(Settings.userPreferences.WikiLocation, transcludeName), releaseTemplateInfo(tm).ToString());


            if (options["cBwikiPreview"])
            {
                reportAll(sb.ToString(), sender);
                reportProgressText(ErrorType.general,"\r\n", sender);
            }
            tocDone = true;
            reportProgressText(ErrorType.information, "Aangemaakt bestand " + Path.GetFileName(fileName) + "\r\n", sender);
            reportProgressText(ErrorType.information, "Aangemaakt bestand " + Path.GetFileName(transcludeName) + "\r\n", sender);
            reportProgressText(ErrorType.general, "== Einde aanmaken inhoudsopgave en issuepagina\r\n", sender);
            return true;
        }

        /// <summary>
        /// Haalt informatie van alle (pre)publicaties in de huidige publicatietaal.
        /// Met deze informatie wordt de tekst voor een wiki transclude page aangemaakt waarmee aangegeven kan worden of de huidige publicatie de meest recente is.
        /// 23-05-23: Transclude file gewijzigd om ook aan te geven wat de laatste release van de zibs is ipv alleen de laatste prerelease (aanleiding zib issue ....)
        /// </summary>
        /// <returns>StringBuilder met wiki tekst voor transclude pagina.</returns>
        private StringBuilder releaseTemplateInfo(textManager tm)
        {
            StringBuilder transcludeText = new StringBuilder();
            List<string[]> allReleases = Settings.XGenConfig.listReleases();
            if (allReleases.Count > 0)
            {
                string lastRelease = allReleases.Where(x => x[2] == "0").Max(y => y[0]);  //17-05-23 experimental
                string lastPreRelease = allReleases.Max(x => x[0]);
                string lastPreReleasePage = tm.getWikiLabel("wikiReleasePage") + "_" + lastPreRelease + "(" + Settings.zibcontext.pubLanguage + ")";
                string lastReleasePage = tm.getWikiLabel("wikiReleasePage") + "_" + lastRelease + "(" + Settings.zibcontext.pubLanguage + ")";
                transcludeText.AppendLine("<!-- Dit template is onderdeel van de ZIB wiki. Het bevat informatie over de publicatie versies -->");

                if (lastPreRelease != lastRelease)
                {

                    transcludeText.AppendLine("{{#ifeq:{{{1}}}|" + lastPreReleasePage);
                    transcludeText.AppendLine("|<div align = center>" + tm.getWikiLabel("tcPreReleaseWarning") + " " + string.Format(tm.getWikiLabel("tcLastReleasePage"), lastReleasePage, lastRelease) + "</div>");
                    transcludeText.AppendLine("|{{#ifeq:{{{1}}}|" + lastReleasePage);
                    transcludeText.AppendLine("|<div align = center>" + tm.getWikiLabel("tcLastReleaseInfo") + "</div>");
                    transcludeText.AppendLine("<div align = center>" + string.Format(tm.getWikiLabel("tcLastPreReleasePage"), lastPreReleasePage, lastPreRelease) + "</div>");
                    transcludeText.AppendLine("|<div align = center>" + tm.getWikiLabel("tcLastPreReleaseWarning") + "</div>");
                    transcludeText.AppendLine("<div align = center>" + string.Format(tm.getWikiLabel("tcLastPreReleasePage"), lastPreReleasePage, lastPreRelease) + " " + string.Format(tm.getWikiLabel("tcLastReleasePage"), lastReleasePage, lastRelease) + "</div>|}}|}}");
                }
                else
                {
                    transcludeText.AppendLine("{{#ifeq:{{{1}}}|" + lastReleasePage + "||<div align = center>" + tm.getWikiLabel("tcLastPreReleaseWarning") + "</div>");
                    transcludeText.AppendLine("<div align = center>" + string.Format(tm.getWikiLabel("tcLastReleasePage"), lastReleasePage, lastRelease) + "</div>|}}");
                }
            }
            else
            {
                transcludeText.AppendLine("---");
            }
            return transcludeText;
        }

        /// <summary>
        /// Method buttonActionError
        /// ActionError wordt aangeroepen als overwachts een onbekende knop doorgegeven wordt,
        /// waar geen action aan gekoppeld is.
        /// </summary>
        /// <returns>Bool: vaste waarde false</returns>

        private bool buttonActionError(Dictionary<string, bool> options, object sender)
        {
            foreach (var _radioButton in gBAction.Controls.OfType<RadioButton>())
                if (_radioButton.Checked)
                {
                    reportProgressText(ErrorType.warning, "================= Fout!!!! =====================\r\n", sender);
                    reportProgressText(ErrorType.warning, " Aan de knop " + _radioButton.Name + " (" + _radioButton.Text + ") is geen actie gekoppeld\r\n", sender);
                }
            return false;
        }



        /// <summary>
        /// Method actionLanguageCheck
        /// LanguageCheck is een action hulp methode, die aangeeft of een action uitgevoerd kan worden 
        /// voor de taal status (single of multi) van de gekozen zib.
        /// </summary>
        /// <param name="action">String: Action waarvoor de taal gecheckt wordt, zoals wiki, docx, xls etc.</param>
        /// <returns>boolean die aangeeft of de actie uitgevoerd kan worden.</returns>
        private bool actionLanguageCheck(string action, object worker)
        {
            if (action == "wiki" || action == "zrtf" || action == "xls")
            {
                if (zib.Language == textLanguage.Multi)
                {
                    MessageBox.Show("Deze actie kan alleen uitgevoerd worden voor ééntalige Zib's", "Let op!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                if (zib.getZibPublishLanguage() != Settings.zibcontext.pubLanguage)
                {
                    reportProgressText(ErrorType.warning, "Bouwsteentaal " + zib.getZibPublishLanguage().ToString() + " van " + zib.Name + " en de gekozen publicatietaal " + Settings.zibcontext.pubLanguage.ToString() + " komen niet overeen\r\n", worker);
                    reportProgressText(ErrorType.warning, "Opdracht wordt niet uitgevoerd\r\n", worker);
                    return false;
                }
            }
            else if (action == "register" || action == "references")
            {
                if (zib.Language != textLanguage.Multi)
                {
                    DialogResult answ = MessageBox.Show("Dit is een single language Zib. Doorgaan?", "Waarschuwing",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (answ != DialogResult.Yes)
                        return false;
                    else
                        return true;
                }
            }
            else if (action == "xmi")
            {
                if (zib.Language != textLanguage.Multi)
                {
                    MessageBox.Show("Dit is een single language Zib. Er kan geen XML aangemaakt worden");
                    return false;
                }

            }
            return true;
        }


        private void copyFiles(string baseDirectory, object worker)
        {
            string target;
            bool applyToAll = false;
            bool overwrite = false;
            reportStatus("Saving results...", worker);
            reportProgressText(ErrorType.general, "\r\n== Bestanden kopieeren naar " + tbResultLocation.Text + "\r\n", worker);
            if (!Directory.Exists(baseDirectory)) Directory.CreateDirectory(baseDirectory);
            if (!Directory.Exists(getStandardDirectory("XML", baseDirectory))) Directory.CreateDirectory(getStandardDirectory("XML", baseDirectory));
            if (!Directory.Exists(getStandardDirectory("XLS", baseDirectory))) Directory.CreateDirectory(getStandardDirectory("XLS", baseDirectory));
            if (!Directory.Exists(getStandardDirectory("Wiki", baseDirectory))) Directory.CreateDirectory(getStandardDirectory("Wiki", baseDirectory));
            if (!Directory.Exists(getStandardDirectory("PNG", baseDirectory))) Directory.CreateDirectory(getStandardDirectory("PNG", baseDirectory));
            if (!Directory.Exists(getStandardDirectory("RTF", baseDirectory))) Directory.CreateDirectory(getStandardDirectory("RTF", baseDirectory));
            foreach (var directory in Directory.GetDirectories(Settings.userPreferences.sessionBase))
            {
                if (Directory.GetFiles(directory).Length > 0)
                {
                    string targetDir = Path.Combine(baseDirectory, Path.GetFileName(directory));
                    foreach (var file in Directory.GetFiles(directory))
                    {
                        target = Path.Combine(targetDir, Path.GetFileName(file));
                        if (File.Exists(target))
                        {
                            if (!overwrite && !applyToAll)
                            {
                                var copyDialog = new frmCopyChoice();
                                copyDialog.FileName = Path.GetFileName(file);
                                copyDialog.ShowDialog();
                                if (copyDialog.Overwrite) overwrite = true;
                                if (copyDialog.ApplyToAll) applyToAll = true;
                                copyDialog.Dispose();
                            }
                            if (overwrite) File.Copy(file, target, overwrite);
                        }
                        else
                        {
                            overwrite = false;
                            File.Copy(file, target);
                        }
                        if (!applyToAll) overwrite = false;
                    }
                }
            }
            reportProgressText(ErrorType.general, "== Bestanden gekopieerd\r\n", worker);
            reportStatus("Ready", worker);

        }

        private void Zib_NewMessage(object sender, ErrorMessageEventArgs e)
        {
            reportProgressText(e.MessageType, e.MessageText + "\r\n", zibExtractor);
        }


        private bool IsInRelease(EA.Package eaZib, string releaseDescription, textLanguage language)
        {
            List<string> zibVersions;
            List<string[]> zibsInRelease = Settings.XGenConfig.getZibsInRelease(releaseDescription, language);
            if (zibsInRelease.Count > 0)
            {
                zibVersions = zibsInRelease.Where(x => x[1] == zibName.shortName(eaZib.Name))?.Select(y => y[2]).ToList();
            }
            else
                return false;
            return zibVersions.Count == 0 ? false : zibVersions.Contains(zibName.Version(eaZib.Name));
        }



        private void rePopulateProjectView(object device)
        {
            ((BackgroundWorker)device).ReportProgress(0, new reportPars { rePopulate = true });
        }

        private void reportProgressText(ErrorType type, string message, object device)
        {
            ((BackgroundWorker)device).ReportProgress(0, new reportPars { informationMessage = new InformationMessage(type, message), addLine = true });
        }
        private void reportAll(string message, object device)
        {
            ((BackgroundWorker)device).ReportProgress(0, new reportPars { informationMessage = new InformationMessage(ErrorType.general, message), addLine = false });
        }
        private void reportStatus(string message, object device)
        {
            ((BackgroundWorker)device).ReportProgress(0, new reportPars { state = message });
        }
        private void reportError(string message, object device)
        {
            ((BackgroundWorker)device).ReportProgress(0, new reportPars { errorMessage = message });
        }
        private void reportBarLabel(string message, object device)
        {
            ((BackgroundWorker)device).ReportProgress(0, new reportPars { labelMessage = message });
        }

        private void reportProgressBar(int n , object device)
        {
            ((BackgroundWorker)device).ReportProgress(n);
        }


        /// <summary>
        /// Class voor de input object parameter van background worker
        /// </summary>

        class zibExtractorPars
        {
            public List<zibAction> actionList;
            public Dictionary<string, bool> options;
            public string copyPath;
            public ErrorType logLevel;
        }

        /// <summary>
        /// Class voor de report userState object parameter van background worker
        /// </summary>

        class reportPars
        {
            public string state;
            public bool rePopulate;
            public string errorMessage;
            public bool addLine;
            public InformationMessage informationMessage;
            public string labelMessage;
        }

        struct InformationMessage
        {
            public ErrorType Type { get; set; }
            public string Message { get; set; }

            public InformationMessage(ErrorType type, string message)
            {
                Type = type;
                Message = message;
            }

        }

    }
}
