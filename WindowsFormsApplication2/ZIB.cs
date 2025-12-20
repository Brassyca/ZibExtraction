using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Text;
using System.Drawing;
using System.Diagnostics;
using System.Net;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Zibs.Configuration;
using Zibs.ExtensionClasses;
using Zibs.CodeCheck;
using Zibs.tekstManager;
using Zibs.Valuesets;
using XConfClasses;




namespace Zibs
{
    namespace ZibExtraction
    {

        public partial class ZIB
        {
            public textManager tm;
            public int maxIndent = 5;

            const string NLCM =    @"2.16.840.1.113883.2.4.3.11.60.40.3";

            EA.Package zp;
            string zFullname, zName, zShortname, zShortnameEN, zVersion, zPrefix, zOID, zStatus, zPublishDate, zPublicationStatus, zPublicationDate;
            textLanguage currentLanguage;
            List<concept> conceptList = new List<concept>();
            concept root = new concept();
            Func<string, string> getIssueSummary;
            readonly EA.Repository r;

            int translateFromYear;


            // =======================================================
            // Event voor het rapporteren van status- en foutmeldingen
            // =======================================================

            public delegate void NewMessageEventHandler(object sender, ErrorMessageEventArgs e);
            public event NewMessageEventHandler NewMessage;

            protected virtual void OnNewMessage(ErrorMessageEventArgs e)
            {
                NewMessage?.Invoke(this, e);
            }


            public string Fullname
            {
                // Volledige bouwsteennaam met prefix en versie
                get { return zFullname; }
                set { zFullname = value; }
            }

            public string Shortname
            {
                // Bouwsteennaam zonder prefix en versie
                get { return zShortname; }
                set { zShortname = value; }
            }
            public string ShortnameEN
            {
                // Bouwsteennaam zonder prefix en versie
                get { return zShortname; }
                set { zShortname = value; }
            }

            public string Name
            {
                // Bouwsteennaam zonder versie
                get { return zName; }
                set { zName = value; }
            }

            public string Prefix
            {
                get { return zPrefix; }
                set { zPrefix = value; }
            }

            public string Version
            {
                get { return zVersion; }
                set { zVersion = value; }
            }

            public string Status
            {
                get { return zStatus; }
                set { zStatus = value; }
            }
            public string PublicationStatus
            {
                get { return zPublicationStatus; }
                set { zPublicationStatus = value; }
            }
            public string PublicationDate
            {
                get { return zPublicationDate; }
                set { zPublicationDate = value; }
            }

            public string ZibOID
            {
                get { return zOID; }
                set { zOID = value; }
            }

            public string ZibPDate
            {
                get { return zPublishDate; }
                set { zPublishDate = value; }
            }

            public textLanguage Language
            {
                get { return currentLanguage; }
                set { currentLanguage = value; }
            }

            public EA.Package EApackage
            {
                get { return zp; }
            }

            //public ZIB(EA.Package z, EA.Repository repos, Func<string, string> getText, Word.Application apWord, Excel.Application excel)
            public ZIB(EA.Package z, EA.Repository repos, Func<string, string> getText)
            {

                zp = z;
                r = repos;
                getIssueSummary = getText;
                zFullname = zp.Name;
                zShortname = zibName.shortName(zFullname);
                zShortnameEN = zibName.shortName(zp.Alias.Replace("EN:", "").Trim());
                zName = zibName.Name(zFullname);
                zVersion = zibName.Version(zFullname);
                zPrefix = zibName.Prefix(zFullname);
                currentLanguage = getZibPublishLanguage();
                zStatus = getZibLifeCycleStatus();
                zPublicationStatus = getZibPublicationStatus();
                zPublicationDate = getZibPublicationDate();
                getZibOID();
            }



            public string getPublishDate()
            {
                foreach (EA.TaggedValue t in zp.Element.TaggedValuesEx)
                {
                    if (t.Name == "DCM::PublicationDate") return t.Value;
                }
                return "";
            }

            public textLanguage getZibPublishLanguage()
            {
                return ZIB.getZibPublishLanguage(zp);
            }
            public static textLanguage getZibPublishLanguage(EA.Package zibPackage)
            {
                textLanguage publishLanguage = textLanguage.Multi;
                EA.TaggedValue t = zibPackage.Element.TaggedValuesEx.GetByName("HCIM::PublicationLanguage");
                if (t != null)
                    if (!Enum.TryParse(t.Value, out publishLanguage)) publishLanguage = textLanguage.Multi;
                return publishLanguage;
            }

            public void getZibOID()
            {
                EA.TaggedValue t = zp.Element.TaggedValuesEx.GetByName("DCM::Id");
                if (t != null)
                    ZibOID = t.Value;
                else
                    ZibOID = "Unknown";
            }


            public void setZibPublishLanguage(textLanguage publishLanguage)
            {
                EA.TaggedValue t;
                bool succes = false;
                t = zp.Element.TaggedValuesEx.GetByName("HCIM::PublicationLanguage");
                if (t == null)
                {
                    t = zp.Element.TaggedValues.AddNew("HCIM::PublicationLanguage", "string");
                    t.Name = "HCIM::PublicationLanguage";
                }
                t.Value = publishLanguage.ToString();
                succes = t.Update();
                zp.Element.TaggedValuesEx.Refresh();
            }

            public void getSections(ref List<string> temp, outputType o, textLanguage language)
            {
                var funcdict = new Dictionary<sectionType, Func<sectionType, string>>
            {
                { sectionType.Header,  (s) => header2text(s.ToString().Replace('_', ' '),o,language) },
                { sectionType.ZIBtags,  (s) => ZIBtags2text(s.ToString().Replace('_', ' '),o,language) },
                { sectionType.Revision_History,  (s) => revHist2text(s.ToString().Replace('_', ' '),o,language) },
                { sectionType.Concept, (s) => section2text(s.ToString().Replace('_', ' '),o,language) },
                { sectionType.Purpose, (s) => section2text(s.ToString().Replace('_', ' '),o,language) },
                { sectionType.Evidence_Base, (s) => section2text(s.ToString().Replace('_', ' '),o,language) },
                { sectionType.Patient_Population, (s) => section2text(s.ToString().Replace('_', ' '),o,language) },
                { sectionType.Information_Model, (s) => informationModel2text(s.ToString().Replace('_', ' '),o,language) },
                { sectionType.Example_Instances, (s) => example2text(s.ToString().Replace('_', ' '),o,language) },
                { sectionType.Instructions, (s) => section2text(s.ToString().Replace('_', ' '),o,language) },
                { sectionType.Issues, (s) => section2text(s.ToString().Replace('_', ' '),o,language) },
                { sectionType.References, (s) => refs2text(s.ToString().Replace('_', ' '),o,language) },
                { sectionType.Traceability_to_other_Standards, (s) => section2text(s.ToString().Replace('_', ' '),o,language) },
                { sectionType.Disclaimer, (s) => termsofuse2text(s.ToString().Replace('_', ' '),o,language) },
                { sectionType.Terms_of_Use, (s) => termsofuse2text(s.ToString().Replace('_', ' '),o,language) },
                { sectionType.Copyrights, (s) => termsofuse2text(s.ToString().Replace('_', ' '),o,language) },
                { sectionType.Assigning_Authorities, (s) => aa2text(sectionType.Information_Model.ToString().Replace('_', ' '),o,language) },
                { sectionType.Valuesets, (s) => valueset2text(sectionType.Information_Model.ToString().Replace('_', ' '),o,language) },
                { sectionType.Footer, (s) => footer2text(s.ToString().Replace('_', ' '),o,language) }
            };
                if (o == outputType.wiki || o == outputType.xls)
                {
                    tm = new textManager("ZibExtractionLabels.cfg");
                    if (tm.dictionaryOK == false)
                    {
                        OnNewMessage(new ErrorMessageEventArgs(ErrorType.warning, "getSections: Geen labeldictionary beschikbaar"));
                        return;
                    }
                    tm.Language = Settings.zibcontext.pubLanguage;
                }

                foreach (var value in Enum.GetValues(typeof(sectionType)))
                {
                    string result = funcdict[(sectionType)value]((sectionType)value);
                    if (!string.IsNullOrEmpty(result)) temp.Add(result);
                    OnNewMessage(new ErrorMessageEventArgs(ErrorType.information, "Sectie: " + value.ToString()));
                }

                StringBuilder sb = new StringBuilder();
                string fileName, transcludeName;
                foreach (string s in temp)
                    sb.AppendLine(s);
                fileName = zibName.wikiLink(zFullname, currentLanguage, !Settings.zibcontext.zibPrefix.Contains("template"));
                if (o == outputType.wiki)
                {
                    fileName += @".wiki";
                    File.WriteAllText(Path.Combine(Settings.userPreferences.WikiLocation, fileName), sb.ToString());

                    // Schrijf de transclude pagina met versie informatie. Voorlopig niet voor blauwdrukken
                    if (!Settings.zibcontext.zibPrefix.Contains("template"))
                    {
                        transcludeName = getTranscludeFilename(Settings.zibcontext.pubLanguage) + ".template";
                        File.WriteAllText(Path.Combine(Settings.userPreferences.WikiLocation, transcludeName), releaseInfo().ToString());
                    }
                }
                else if (o == outputType.text) //was wiki
                {
                    fileName += @".txt";
                    File.WriteAllText(Path.Combine(Settings.userPreferences.WikiLocation, fileName), sb.ToString());
                }
                else if (o == outputType.xls)
                {
                    fileName = zibName.fileName2(zFullname, currentLanguage) + ".xlsx";
                    saveToSpeadsheet(Path.Combine(Settings.userPreferences.XLSLocation, fileName), sb.ToString());
                }

            }


            /// <summary>
            /// Haalt informatie van alle versies van de huidige bouwsteen op met de naam in de huidige publicatietaal.
            /// Met deze informatie wordt de tekst voor een wiki transclude page aangemaakt waarmee aangegeven kan worden of de huidige versie de meest recente is.
            /// Tevens wordt een lijst gemaakt van alle beschikbare versies en de publicaties waarin deze voorkomen. Deze wordt ook als transclude op de wiki pagina's
            /// getoond. Door dit op een transclude pagina te zetten kunnen eerder gepubliceerde geupdate worden zonder ze opnieuw te genereren.
            /// </summary>
            /// <returns>StringBuilder lijst met wiki tekst voor transclude pagina.</returns>
            private StringBuilder releaseInfo()
            {
                StringBuilder transcludeText = new StringBuilder();
                List<string[]> allReleases = Settings.XGenConfig.GetAllReleases(zOID, Settings.zibcontext.pubLanguage);
                if (allReleases.Count > 0)
                {
                    string[] lastVersionInfo = allReleases[allReleases.Count - 1];
                    string lastVersion = zibName.wikiLink(lastVersionInfo[3] + "-v" + lastVersionInfo[2], Settings.zibcontext.pubLanguage, lastVersionInfo[0]);
                    transcludeText.AppendLine("<!-- Dit template is onderdeel van de ZIB wiki. Het bevat informatie over de versies van zib " + zOID + ". -->");
                    transcludeText.AppendLine("{{#ifeq:{{{1}}}|1|");
                    transcludeText.AppendLine("{{#ifeq:{{{2}}}|" + lastVersion + "||<div align = center>" + String.Format(tm.getWikiLabel("tcLastRelease"), lastVersion) +  "</div>|}}");
                    transcludeText.AppendLine("}}");
                    transcludeText.AppendLine("{{#ifeq:{{{1}}}|2|");
                    transcludeText.AppendLine("<ul>");
                    //27-09-2021 Ondescheid tussen publicatie en pre-publicatie toegevoegd incl prepublicatie volgnummer
                    foreach (var _release in allReleases)
                    {
                        transcludeText.AppendLine("{{#ifeq:{{{2}}}|" + _release[0] + "||<li>[[" + zibName.wikiLink(_release[3] + "-v" + _release[2], Settings.zibcontext.pubLanguage, _release[0]) +
                            " | " + (_release[1] == "0" ? (tm.getWikiLabel("hdPublication") + " " + _release[0]) : (tm.getWikiLabel("hdPrepublication") + " " + _release[0]) + "-" + _release[1]) + 
                            ", (" + tm.getWikiLabel("hdVersion") + " " + _release[2] + ")]]</li>}}");
                    }
                    transcludeText.AppendLine("</ul>|}}");
                }
                else
                {
                    transcludeText.AppendLine("---");
                }

                return transcludeText;

            }



            // ===========================
            // textfuncties per sectietype 
            // ===========================

            private string header2text(string sectionName, outputType mode, textLanguage language)
            {
                StringBuilder headerText = new StringBuilder();
                if (mode == outputType.wiki)
                {
                    if (!Settings.zibcontext.zibPrefix.Contains("template"))
                    {
                        headerText.AppendLine("<!-- Hieronder wordt een transclude page aangeroepen -->");
                        headerText.AppendLine("{{" + getTranscludeFilename(Settings.zibcontext.pubLanguage) + "|1|" + zibName.wikiLink(zFullname, Settings.zibcontext.pubLanguage) + "}}");
                        headerText.AppendLine("<!-- Tot hier de transclude page -->");
                    }

                    headerText.AppendLine("==" + tm.getWikiLabel("hdGeneralInformation") + "==");
                    headerText.Append(tm.getWikiLabel("hdName") + ": '''" + zName + "'''");
                    foreach (textLanguage value in Enum.GetValues(typeof(textLanguage)))
                        if (value != textLanguage.Multi && value != Settings.zibcontext.pubLanguage)
                        {
                            string _zibname = Settings.XGenConfig.getLanguageSpecificName(zOID, value, zVersion);
                            headerText.Append(" [[Bestand:" + value + ".png|link=" + zibName.wikiLink(_zibname + "-v" +zVersion, value,!Settings.zibcontext.zibPrefix.Contains("template")) + "]]");
                        }
                    headerText.AppendLine("<BR>");
                    headerText.AppendLine(tm.getWikiLabel("hdVersion") + ": '''" + zVersion + "''' <br>");
                    headerText.AppendLine(tm.getWikiLabel("hdStatus") + ":" + zStatus + "<br>");
                    if (!Settings.zibcontext.zibPrefix.Contains("template")) headerText.AppendLine(tm.getWikiLabel("hdPublication") + ": '''" + (Settings.zibcontext.publicatie ?? "") + "''' <br>");
                    //headerText.AppendLine(tm.getWikiLabel("hdPublication") + ": '''[[" + Settings.zibcontext.zibCategory + ": publicatieinformatie|" + (Settings.zibcontext.publicatie ?? "") + "]]''' <br>");
                    //headerText.AppendLine(tm.getWikiLabel("hdWikipageStatus") + ": '''" + Settings.wikicontext.wikiPagesStatus + "'''<BR>");
                    headerText.AppendLine(tm.getWikiLabel("hdPublicationStatus") + ": " + zPublicationStatus + "<br>");
                    headerText.AppendLine(tm.getWikiLabel("hdPublicationDate") + ": " + zPublicationDate);
                    // Hier volgt de Errata transclude pagina
                    headerText.AppendLine("<!-- Aanroep Errata transclude page -->");
                    headerText.AppendLine("{{"+ tm.getWikiLabel("hdErrata") +"|"+ Settings.zibcontext.publicatie + "|{{PAGENAME}}}}");
                    headerText.AppendLine("<!-- tot hier -->");
                    //
                    headerText.AppendLine("-----");
                    if (!Settings.zibcontext.zibPrefix.Contains("template"))
                    {
                        headerText.AppendLine("<div style=\"text-align: right; direction: ltr; margin-left: 1em;\" >[[Bestand: Back 16.png| link= " +
                        tm.getWikiLabel("wikiReleasePage") + "_" + (Settings.zibcontext.publicatie ?? "??") + "(" + Settings.zibcontext.pubLanguage + ")]] " +
                        "[[" + tm.getWikiLabel("wikiReleasePage") + "_" + (Settings.zibcontext.publicatie ?? "??") + "(" + Settings.zibcontext.pubLanguage +
                        ") |" + tm.getWikiLabel("hdBackToMainPage") + " ]]" + "</div>");
                    }
                }
                else if (mode == outputType.text)
                {
                    headerText.AppendLine(tm.getLabel("hdGeneralInformation"));
                    headerText.AppendLine(tm.getLabel("hdName") + ": " + zName);
                    headerText.AppendLine(tm.getLabel("hdVersion") + ": " + zVersion);
                    headerText.AppendLine(tm.getLabel("hdStatus") + ":" + zStatus);
                    headerText.AppendLine(tm.getLabel("hdPublication") + ":" + Settings.zibcontext.publicatie ?? "");
                    headerText.AppendLine(tm.getLabel("hdPublicationStatus") + ": " + zPublicationStatus);
                    headerText.AppendLine(tm.getLabel("hdPublicationDate") + ": " + zPublicationDate);
                }
                else if (mode == outputType.xls)
                {
                    CodeNames cn = new CodeNames();
                    cn.languageRefsetId_EN = Settings.application.Snomed_languageRefsetId_EN;
                    cn.languageRefsetId_NL = Settings.application.Snomed_languageRefsetId_NL;
                    cn.terminologyLink = Settings.application.TerminologyServiceLink;
                    if (int.TryParse(Settings.application.CodeTranslateFromYear, out translateFromYear)) cn.translateFromYear = translateFromYear;

                    string snomedVersion = cn.getCodesystemVersion("SNOMED CT", Settings.zibcontext.publicatie);
                    snomedVersion = snomedVersion.Substring(snomedVersion.IndexOf(':') + 1).Trim();
                    string loincVersion = cn.getCodesystemVersion("LOINC", Settings.zibcontext.publicatie);

                    headerText.AppendLine("<<Sheet>>" + tm.getLabel("xlsAbout") + ";2;2");
                    headerText.AppendLine("<<ColWidth>>15;100");
                    headerText.AppendLine("<<IgnoreNumberAsText>>2;2;7;2");
                    headerText.AppendLine(tm.getLabel("xlsItem") + ";" + tm.getLabel("xlsDescription"));
                    headerText.AppendLine("<<Header>>");
                    headerText.AppendLine(tm.getLabel("xlsName") + ";" + zName);
                    headerText.AppendLine(tm.getLabel("hdVersion") + ";" + zVersion);
                    headerText.AppendLine(tm.getLabel("hdStatus") + ";" + zStatus);
                    if (!Settings.zibcontext.zibPrefix.Contains("template")) headerText.AppendLine(tm.getLabel("hdPublication") + ";" + Settings.zibcontext.publicatie ?? "");
                    headerText.AppendLine(tm.getLabel("hdPublicationStatus") + "; " + zPublicationStatus);
                    headerText.AppendLine(tm.getLabel("hdPublicationDate") + "; " + zPublicationDate);
                    headerText.AppendLine(tm.getLabel("xlsDate") + ";" + DateTime.Now.ToString(new CultureInfo("nl-NL")));
                    headerText.AppendLine(tm.getLabel("xlsInfoBase") + ";" + (Settings.zibcontext.ReleaseInfo +
                        "\r\nSNOMED CT version " + snomedVersion + "\r\nLOINC version " + loincVersion).encodeForXLS());
                    headerText.AppendLine(tm.getLabel("xlsManagedBy") + ";[Hyperlink:https://www.nictiz.nl]Nictiz, Den Haag ");
                    headerText.AppendLine("<<ImageRelative>>NictizLogo("+ currentLanguage + ").png;" + Settings.userPreferences.ImageLocation + ";" + 19 + ";" + 2);
                    headerText.AppendLine("<<ImageRelative>>CreativeCommon.jpg;" + Settings.userPreferences.ImageLocation + ";" + 20 + ";" + 3 + ";" + 400);
                }
                else
                    headerText.Append("");

                return headerText.ToString();
            }

            private string section2text(string sectionName, outputType mode, textLanguage language)
            {
                string sectionText = "";
                EA.Package section = zp.Packages.GetByName(sectionName);
                if (!string.IsNullOrEmpty(section.Notes))
                {
                    if (mode == outputType.wiki)
                    {
                        sectionText = "==" + section.Name + "==\r\n";
                        sectionText += section.Notes.toWiki() + "\r\n";
                    }
                    else if (mode == outputType.text)
                    {
                        sectionText = section.Name + "\r\n";
                        sectionText += section.Notes + "\r\n\r\n";
                    }
                    else if (mode == outputType.xls)
                    {
                        sectionText = ("<<Sheet>>" + tm.getLabel("xlsAbout")) + "\r\n";
                        sectionText += section.Name;
                        string Notes = section.Notes.encodeForXLS();
                        sectionText += ";" + Notes;
                    }
                }
                return sectionText;
            }


            private string termsofuse2text(string sectionName, outputType mode, textLanguage language)
            {
                string touText = "";
                EA.Package section = zp.Packages.GetByName(sectionName);
                if (!string.IsNullOrEmpty(section.Notes))
                {
                    if (mode == outputType.wiki)
                    {
                        touText = "";  // Geen Terms of Use per Wiki Pagina
                    }
                    else if (mode == outputType.text)
                    {
                        touText = section.Name + "\r\n";
                        touText += section.Notes + "\r\n\r\n";
                    }
                    else if (mode == outputType.xls)
                    {
                        touText = "<<Sheet>>" + tm.getLabel("xlsTermsOfUse") + ";2;2" + "\r\n";
                        touText += "<<ColWidth>>150" + "\r\n";
                        touText += section.Name + "\r\n";
                        touText += "<<Header>>" + "\r\n";
                        touText += section.Notes.encodeForXLS() + "\r\n";
                        touText += "\r\n";
                    }
                }
                return touText;
            }


            private string ZIBtags2text(string sectionName, outputType mode, textLanguage language)
            {
                StringBuilder zibTagsText = new StringBuilder();
                if (mode == outputType.wiki)
                {
                    zibTagsText.AppendLine("==Metadata==");
                    zibTagsText.AppendLine("{| class=\"wikitable\" style=\"font-size:90%; width: 750px\"");
                    foreach (EA.TaggedValue t in zp.Element.TaggedValuesEx)
                    {
                        zibTagsText.AppendLine("|- ");
                        zibTagsText.AppendLine("|style=\"width:250px; \"|" + t.Name + "||" + t.Value);
                        // vul meteen de zib OID en status variabele
                        if (t.Name == "DCM::Id") zOID = t.Value;
                        if (t.Name == "DCM::Status") zStatus = t.Value;
                    }
                    zibTagsText.AppendLine("|}");
                }
                else if (mode == outputType.text)
                {
                    zibTagsText.AppendLine("Metadata");
                    foreach (EA.TaggedValue t in zp.Element.TaggedValuesEx)
                    {
                        zibTagsText.AppendLine(t.Name + " : " + t.Value);
                        if (t.Name == "DCM::Id") zOID = t.Value;
                    }
                    zibTagsText.AppendLine();
                }
                else if (mode == outputType.xls)
                {
                    zibTagsText.AppendLine("<<Sheet>>Metadata;2;2");
                    zibTagsText.AppendLine("Metadata");
                    zibTagsText.AppendLine("<<Merge>>2");
                    zibTagsText.AppendLine("<<Header>>");

                    zibTagsText.AppendLine("<<ColWidth>>35;70");
                    zibTagsText.AppendLine("<<IgnoreNumberAsText>>2;2;24;2");
                    foreach (EA.TaggedValue t in zp.Element.TaggedValuesEx)
                        zibTagsText.AppendLine(t.Name + ";" + t.Value);
                }
                else
                    zibTagsText.Append("");

                return zibTagsText.ToString();

            }

            private string revHist2text(string sectionName, outputType mode, textLanguage language)
            {
                StringBuilder revHistText = new StringBuilder();
                string srLine;
                bool noIssueDict = false;
                EA.Package section = zp.Packages.GetByName(sectionName);
                if (!string.IsNullOrEmpty(section.Notes))
                {
                    if (!section.Notes.EndsWith(".")) section.Notes += ".";   
                    if (mode == outputType.wiki)
                    {
                        revHistText.AppendLine("==" + section.Name + "==");
                        revHistText.AppendLine("<div class=\"mw-collapsible mw-collapsed\">");
                        if (!string.IsNullOrWhiteSpace(tm.getWikiLabel("noTranslation"))) revHistText.AppendLine("'' " + tm.getWikiLabel("noTranslation") + "''\r\n");
                        using (StringReader reader = new StringReader(section.Notes))
                        {
                            srLine = reader.ReadLine();
                            while (srLine != null)
                            {
                                if (srLine.IndexOf(Settings.bitscontext.issuePrefix) == -1 || noIssueDict)
                                    revHistText.AppendLine(srLine);
                                else
                                {
                                    if (!srLine.EndsWith(".")) srLine += ".";  // Punt wordt vaak vergeten en dan missen we een issue
                                    revHistText.AppendLine("{|");
                                    Regex rgx = new Regex("(ZIB-\\d+)[.,]");
                                    Match m = rgx.Match(srLine);
                                    while (m.Success)
                                    {
                                        string s = m.Groups[1].Captures[0].ToString();
                                        if (getIssueSummary(s) == "Issue summaries niet beschikbaar")
                                        {
                                            OnNewMessage(new ErrorMessageEventArgs(ErrorType.error, "Revision History: Issue summaries niet beschikbaar: genereer eerst de issuelijst."));
                                            revHistText.AppendLine(getIssueSummary(s) + "<BR>");
                                            revHistText.AppendLine(srLine);
                                            noIssueDict = true; // Geen issue informatie aanwezig dus geen de revision history 'as is' 
                                            break;
                                        }

                                        // ingevoegd dd 8-6-2022 om meer en kleinere issuepagina's te krijgen
                                        int page = GetPageNumber(s, Settings.bitscontext.maxIssuesPerPage);
                                        if (page == -1)
                                        {
                                            OnNewMessage(new ErrorMessageEventArgs(ErrorType.error, "Revision History: Issue paginanummer niet te bepalen."));
                                            revHistText.AppendLine(getIssueSummary(s) + "<BR>");
                                            revHistText.AppendLine(srLine);
                                            break;
                                        }
                                        revHistText.AppendLine("|-\r\n|style=\"width:75px; \"|[[ZIBIssues" + Settings.bitscontext.maxIssuesPerPage.ToString() + "_" + page.ToString() + "#" + s + " | " + s + " ]]\r\n|" + getIssueSummary(s));

                                        // tot hier
                                        //revHistText.AppendLine("|-\r\n|style=\"width:75px; \"|[[ZIBIssues#" + s + " | " + s + " ]]\r\n|" + getIssueSummary(s));
                                        m = m.NextMatch();
                                    }

                                    revHistText.AppendLine("|}");
                                }
                                srLine = reader.ReadLine();
                            }
                        }
                        revHistText.AppendLine("</div>");
                    }
                    else if (mode == outputType.text)
                    {
                        revHistText.AppendLine(section.Name);
                        if (!string.IsNullOrWhiteSpace(tm.getLabel("noTranslation"))) revHistText.AppendLine("''" + tm.getLabel("noTranslation") + "''\r\n");
                        revHistText.AppendLine(section.Notes);
                    }
                    else
                    // Geen versie historie in XLS output
                    revHistText.Append("");
                }
                revHistText.Append("");
                return revHistText.ToString();
            }


            private int GetPageNumber(string issueKey, int maxIssuesPerPage)
            {
                int pageNumber;
                if (int.TryParse(issueKey.Replace("ZIB-", "").Trim(), out int issueNo))
                {
                    pageNumber = (int)Math.Floor((issueNo * 1.0f) / maxIssuesPerPage);

                }
                else
                {
                    pageNumber = -1;
                }


                return pageNumber;
            }

            private string informationModel2text(string sectionName, outputType mode, textLanguage language)
            {
                StringBuilder informationModelText = new StringBuilder();
                EA.Package section = zp.Packages.GetByName(sectionName);
                getConceptlist(section);

                if (mode == outputType.wiki)
                {
                    informationModelText.AppendLine("==" + section.Name + "==");
                    informationModelText.AppendLine("<BR>");

                    foreach (KeyValuePair<int, string[]> diagram in getDiagrams(section, language))
                    {
//                        informationModelText.AppendLine("[[Bestand: " + diagram.Value + " | center | link=]]");
                        Dictionary<int, imageMapData> links = getDiagramObjectsLinks(diagram.Key);
                        informationModelText.AppendLine("<imagemap> Bestand:" + diagram.Value[1] + " | center ");
                        foreach (KeyValuePair<int, imageMapData> mapLink in links)
                        {
                            informationModelText.AppendLine("rect " + mapLink.Value.topLeft.X + " " + (-1 * mapLink.Value.topLeft.Y) +
                                 " " + mapLink.Value.bottomRight.X + " " + (-1 * mapLink.Value.bottomRight.Y) + " [[" + mapLink.Value.link +"]]");
                        }
                        informationModelText.AppendLine("desc none");
                        informationModelText.AppendLine("</imagemap>");
                    }
                    informationModelText.AppendLine("<BR>");

                    // Header
                    informationModelText.AppendLine("{| border= \"1\"  width=\"1500px\" style = \"font-size: 9.5pt;  border: solid 1px silver; border-collapse:collapse;\" cellpadding = \"3px\" cellspacing =\"0px\" ");
                    informationModelText.AppendLine("|-style=\"background-color: #1F497D; color: white; font-weight: bold; font-variant: small-caps; \"");
                    informationModelText.Append("|style=\"width:30px;\"|" + tm.getWikiLabel("imType") + "||style=\"width:100px;\"|" + tm.getWikiLabel("imId"));
                    informationModelText.Append("||colspan=\"" + (maxIndent+1).ToString()+ " \" style=\"width:140px;\"|" + tm.getWikiLabel("imConcept") + "||" + tm.getWikiLabel("imCard") + "|");
                    informationModelText.Append("|style=\"width: 600px;\"|" + tm.getWikiLabel("imDefinition") + "||style=\"width:200px;\"|" + tm.getWikiLabel("imDefinitionCode"));
                    informationModelText.AppendLine("||style=\"width:200px;\"|" + tm.getWikiLabel("imReference")); // vervallen ivm alias + " ||" + tm.getWikiLabel("imConstraints") + ""); "||style=\"width:140px;\"|" + tm.getWikiLabel("imAlias") + 
                    informationModelText.Append(root.ToWiki());
                    foreach (concept c in conceptList.OrderByDescending(c => c.packageId).ThenBy(c => c.treepos))
                        informationModelText.Append(c.ToWiki());

                    informationModelText.AppendLine("|}");
                    informationModelText.AppendLine(tm.getWikiLabel("imHover") + "<BR>");
                    informationModelText.AppendLine(tm.getWikiLabel("imLegend") + " [[Bestand:list2.png|link=" + tm.getWikiLabel(Settings.wikicontext.LegendPage) + "]]");
                }
                else if (mode == outputType.text)
                {
                    informationModelText.AppendLine(section.Name);
                    foreach (KeyValuePair<int,string[]> diagram in getDiagrams(section, language))
                        informationModelText.AppendLine("Diagram: " + diagram.Value[1]);
                    informationModelText.AppendLine("Aantal concepten: " + conceptList.Count);
                    //                foreach (concept c in conceptList)
                    informationModelText.AppendLine(root.ToString());
                    foreach (concept c in conceptList.OrderByDescending(c => c.packageId).ThenBy(c => c.treepos))
                        informationModelText.AppendLine(c.ToString());
                    informationModelText.AppendLine("");
                }
                else if (mode == outputType.xls)
                {
                    foreach (KeyValuePair<int, string[]> diagram in getDiagrams(section, language))
                    {
                        informationModelText.AppendLine("<<Sheet>>" + diagram.Value[0]);
                        informationModelText.AppendLine("<<Image>>" + diagram.Value[1] + ";" + Settings.userPreferences.DiagramLocation);
                    }
                    informationModelText.AppendLine("<<Sheet>>Data;2;2");
                    informationModelText.Append(tm.getLabel("imConcept") + ";" + " ;" + " ;" + " ;" + " ;" + " ");
                    informationModelText.Append(";" + tm.getLabel("imAlias") + ";" + tm.getLabel("imType"));
                    informationModelText.Append(";" + tm.getLabel("imCard") + ";" + tm.getLabel("imStereotype") + ";" + tm.getLabel("imId"));
                    informationModelText.Append(";" + tm.getLabel("imDefinition") + ";" + tm.getLabel("imDefinitionCode"));
                    informationModelText.AppendLine(";" + tm.getLabel("imReference") + ";" + tm.getLabel("imConstraints"));
                    informationModelText.AppendLine("<<Header>>");
                    informationModelText.AppendLine("<<ColWidth>>2;2;2;2;2;15;25;5;5;12;15;75;20;30;20");
                    informationModelText.AppendLine("<<IgnoreNumberAsText>>2;9;" + (conceptList.Count()+1).ToString()  + ";9");
                    informationModelText.Append(root.ToXLS());
                    foreach (concept c in conceptList.OrderByDescending(c => c.packageId).ThenBy(c => c.treepos))
                        informationModelText.Append(c.ToXLS());
                }

                return informationModelText.ToString();
            }

            private string example2text(string sectionName, outputType mode, textLanguage language)
            {
                if (Settings.zibcontext.zibPrefix.Contains("template"))
                    return String.Empty;
                // blauwdrukken hebben geen voorbeeldbestand
                string filename;
                StringBuilder exampleText = new StringBuilder();
                EA.Package section = zp.Packages.GetByName(sectionName);

                //              Voorbeeld bestanden zijn er alleen in het Nedelands, patch 12-09
                if (currentLanguage == textLanguage.NL)
                {
                    filename = Path.Combine(Settings.userPreferences.ExampleLocation, zibName.fileName(Fullname, currentLanguage) + @"_Voorbeeld.docx");
                }
                else
                {
                    // maak de Nederlandse naam van de bouwsteen aan
                    string fileNameNL = Settings.zibcontext.zibPrefix + "." + Settings.XGenConfig.getLanguageSpecificName(zOID, textLanguage.NL, zVersion) + "-v" + zVersion.ToString();
                    filename = Path.Combine(Settings.userPreferences.ExampleLocation, zibName.fileName(fileNameNL, textLanguage.NL) + @"_Voorbeeld.docx");
                }

                if (mode == outputType.wiki)
                {
                    exampleText.AppendLine("==" + section.Name + "==");
                    if (!string.IsNullOrWhiteSpace(tm.getLabel("noTranslation"))) exampleText.AppendLine("''" + tm.getWikiLabel("noTranslation") + "''\r\n");
                    zibExample zE = new zibExample();
                    zE.NewMessage += zibexampleMessage;

                    exampleText.Append(zE.ReadContent(filename));

                    // 24-05-2023 Toegevoegd om opmerkingen bij het voorbeeld te kunnen plaatsen
                    if (!string.IsNullOrEmpty(section.Notes))
                    {
                        exampleText.Append("\r\n" + section.Notes.toWiki() + "\r\n");
                    }

                    // Voeg ook de notes van het diagram toe, want dat zal toch wel fout gedaan worden.

                    EA.Diagram exampleDiagram = section.Diagrams.GetAt(0);
                    if (!string.IsNullOrEmpty(exampleDiagram.Notes))
                    {
                        exampleText.Append((!string.IsNullOrEmpty(section.Notes)? "" : "<BR>\r\n") + exampleDiagram.Notes.toWiki() + "\r\n");
                    }
                }
                else if (mode == outputType.text)
                {
                    exampleText.AppendLine(section.Name);
                    if (!string.IsNullOrWhiteSpace(tm.getLabel("noTranslation"))) exampleText.AppendLine("''" + tm.getLabel("noTranslation") + "''\r\n");
                    exampleText.AppendLine(tm.getLabel("exNoRepresentation") + " " + filename);
                }
                else
                    exampleText.Append("");
                return exampleText.ToString();
            }

            private string refs2text(string sectionName, outputType mode, textLanguage language)
            {
                string refsText = "";
                int sChar, eChar;
                EA.Package section = zp.Packages.GetByName(sectionName);
                if (!string.IsNullOrEmpty(section.Notes))
                {
                    if (mode == outputType.wiki)
                    {
                        refsText = "==" + section.Name + "==\r\n";
                        sChar = section.Notes.IndexOf("<a href");
                        int l = section.Notes.Length;
                        while (sChar != -1)
                        {
                            eChar = section.Notes.IndexOf("\">", sChar + 7);
                            section.Notes = section.Notes.Substring(0, sChar - 1) + section.Notes.Substring(eChar + 2);
                            eChar = section.Notes.IndexOf("</a>", 0) + 4;
                            sChar = section.Notes.IndexOf("<a href", eChar);
                        }
                        section.Notes = section.Notes.Replace("</a>", "");
                        refsText += section.Notes + "\r\n";
                    }
                    else if (mode == outputType.text)
                    {
                        refsText = section.Name + "\r\n";
                        refsText += section.Notes + "\r\n\r\n";
                    }
                }
                else
                    // Geen referenties in XLS output
                    refsText = "";

                return refsText;
            }

            private string aa2text(string sectionName, outputType mode, textLanguage language)
            {
                string aaText = "";
                string temp;
                if (mode == outputType.wiki)
                {
                    aaText += "==" + tm.getWikiLabel("aaAssigningAuthorities") + "==\r\n";
                    aaText += tm.getWikiLabel("aaHeader") + " <BR>\r\n";
                }
                else if (mode == outputType.text)
                {
                    aaText += tm.getLabel("aaAssigningAuthorities") + "\r\n";
                }
                EA.Package section = zp.Packages.GetByName(sectionName);
                temp = createAssigningAuthorities(section, mode);
                aaText = string.IsNullOrEmpty(temp) ? "" : aaText + temp;
                return aaText;
            }

            private string valueset2text(string sectionName, outputType mode, textLanguage language)
            {
                string valuesetText = "";
                string temp;
                if (mode == outputType.wiki)
                    valuesetText += "== Valuesets ==\r\n";
                else if (mode == outputType.text)
                    valuesetText += "Valuesets\r\n";

                EA.Package section = zp.Packages.GetByName(sectionName);
                temp = getValueSets(section, mode);
                valuesetText = string.IsNullOrEmpty(temp) ? "" : valuesetText + temp;

                return valuesetText;
            }

            private string footer2text(string sectionName, outputType mode, textLanguage language)
            {
                StringBuilder footerText = new StringBuilder();
                List<string[]> zibReferences;
                CodeNames cn = new CodeNames();
                cn.languageRefsetId_EN = Settings.application.Snomed_languageRefsetId_EN;
                cn.languageRefsetId_NL = Settings.application.Snomed_languageRefsetId_NL;
                cn.terminologyLink = Settings.application.TerminologyServiceLink;
                if (int.TryParse(Settings.application.CodeTranslateFromYear, out translateFromYear)) cn.translateFromYear = translateFromYear;

                string snomedVersion = cn.getCodesystemVersion("SNOMED CT", Settings.zibcontext.publicatie);
                string loincVersion = cn.getCodesystemVersion("LOINC", Settings.zibcontext.publicatie);
                if (mode == outputType.wiki)
                {
                    if (!Settings.zibcontext.zibPrefix.Contains("template"))
                    {
                        footerText.AppendLine("==" + tm.getWikiLabel("ftOtherReleases") + "==");
                        // ---
                        footerText.AppendLine("<!-- Hieronder wordt een transclude page aangeroepen -->");
                        footerText.AppendLine("{{" + getTranscludeFilename(Settings.zibcontext.pubLanguage) + "|2|" + Settings.zibcontext.publicatie + "}}");
                        footerText.AppendLine("<!-- Tot hier de transclude page -->");
                        // ---

                        footerText.AppendLine("==" + tm.getWikiLabel("ftReferences") + "==");
                        footerText.AppendLine("====" + tm.getWikiLabel("ftRefersTo") + "====");
                        zibReferences = Settings.XGenConfig.UsesZibs(zOID, Settings.zibcontext.publicatie, Settings.zibcontext.PreReleaseNumber.ToString(), Settings.zibcontext.pubLanguage);
                        if (zibReferences.Count > 0)
                        {
                            foreach (string[] zibRef in zibReferences)
                                footerText.AppendLine("*[[" + zibName.wikiLink(zibRef[0] + "-v" + zibRef[1], Settings.zibcontext.pubLanguage) + "|" + zibRef[0] + "-v" + zibRef[1] + "]]");
                        }
                        else
                            footerText.AppendLine(":--");

                        footerText.AppendLine("====" + tm.getWikiLabel("ftReferredBy") + "====");
                        zibReferences = Settings.XGenConfig.UsedInZibs(zOID, Settings.zibcontext.publicatie, Settings.zibcontext.PreReleaseNumber.ToString(), Settings.zibcontext.pubLanguage);
                        if (zibReferences.Count > 0)
                        {
                            foreach (string[] zibRef in zibReferences)
                                footerText.AppendLine("*[[" + zibName.wikiLink(zibRef[0] + "-v" + zibRef[1], Settings.zibcontext.pubLanguage) + "|" + zibRef[0] + "-v" + zibRef[1] + "]]");
                        }
                        else
                            footerText.AppendLine(":--");


                        footerText.AppendLine("==" + tm.getWikiLabel("ftHeader") + "==");

                        footerText.AppendLine(tm.getWikiLabel("ftReferenceIntro"));
                        footerText.AppendLine("<ul>");

                        /* De Art-Decor url wordt uit een sjabloon gehaald: ArtDecorLinks
                        DateTime result;
                        string dateTime = DateTime.TryParse(getPublishDate(), out result) ? result.ToString("s") : "0000-00-00T00:00:00";
                        footerText.AppendLine("<li>" + tm.getWikiLabel("ftArtDecorReference") + " [" + Settings.wikicontext.ArtDecorUrl + "/art-decor/decor-scenarios--" + Settings.wikicontext.ArtDecorRepository + "-?id=" + zOID.Replace(NLCM, Settings.wikicontext.ArtDecorProjectOID) +
                        "&effectiveDate=" + dateTime + "&language=" + multilanguageText.languageCode[Settings.zibcontext.pubLanguage] + "&scenariotree=false [[File:artdecor.jpg|16px|link=]]]</li>");
                        */
                        footerText.AppendLine("<li>" + tm.getWikiLabel("ftArtDecorReference") + " {{ArtDecorLinks|" + Settings.zibcontext.publicatie + "|" + zOID.Replace((NLCM + "." ), "") +  "}}</li>");

                        //                  De FHIR url wordt nu uit een sjabloon (SimplifierLinks) pagina gehaald
                        footerText.AppendLine("<li>" + tm.getWikiLabel("ftSimplifierReference") + " {{SimplefierLinks|" + Settings.zibcontext.publicatie + "|"+ zShortnameEN + "}}</li>");
                        //                    footerText.AppendLine("<li>" + tm.getWikiLabel("ftSimplifierReference") + " [" + Settings.wikicontext.SimplifierUrl + " [[File:fhir.png|link=]]]</li>");
                        footerText.AppendLine("</ul>");

                        footerText.AppendLine("==" + tm.getWikiLabel("ftDownloadTitle") + "==");
                        string filePostfix = (Settings.zibcontext.publicatie ?? "") + (language == textLanguage.Multi ? "" : language.ToString());
                        filePostfix = filePostfix.Length > 0 ? "(" + filePostfix + ")" : filePostfix;
                        string xlsFilename = zFullname + filePostfix + ".xlsx";
                        string pdfFilename = zFullname + filePostfix + ".pdf";
                        footerText.AppendLine(String.Format(tm.getWikiLabel("ftDownloads"), pdfFilename, xlsFilename));
                    }
                    footerText.AppendLine("==" + tm.getWikiLabel("ftHeader2") + "==");
                    footerText.AppendLine(tm.getWikiLabel("ftInfoBase") + " " + Settings.XGenConfig.getReleaseInfo(Settings.zibcontext.publicatie, Settings.zibcontext.PreReleaseNumber.ToString(), Settings.zibcontext.pubLanguage) + " <BR>");
                    //                  Toevoegen codesysteem versies
                    if (snomedVersion != "" && loincVersion != "")
                    {
                        footerText.AppendLine(tm.getLabel("ftCodeSystemReleases") + ":");
                        footerText.AppendLine("<ul>");
                        footerText.AppendLine("<li>" + (snomedVersion == "" ? "SNOMED --" : snomedVersion) + "</li>");
                        footerText.AppendLine("<li>LOINC version " + loincVersion + "</li>");
                        footerText.AppendLine("</ul>");
                    }
                    footerText.AppendLine(tm.getWikiLabel("ftConditions") + " [[Bestand:list2.png|link=" + tm.getWikiLabel(Settings.wikicontext.MainPage) + "]]<BR>");
                    footerText.AppendLine(string.Format(tm.getWikiLabel("ftDate"),
                            DateTime.Now.ToString(new CultureInfo("en-GB")),
                            typeof(ZibExtraction).Assembly.GetName().Version) + " <BR>");
                    footerText.AppendLine("-----");
                    if (!Settings.zibcontext.zibPrefix.Contains("template"))
                    {
                        footerText.AppendLine("<div style=\"text-align: right; direction: ltr; margin-left: 1em;\" >[[Bestand: Back 16.png| link= " +
                        tm.getWikiLabel("wikiReleasePage") + "_" + (Settings.zibcontext.publicatie ?? "??") + "(" + Settings.zibcontext.pubLanguage + ")]] " +
                        "[[" + tm.getWikiLabel("wikiReleasePage") + "_" + (Settings.zibcontext.publicatie ?? "??") + "(" + Settings.zibcontext.pubLanguage +
                        ") |" + tm.getWikiLabel("hdBackToMainPage") + " ]]" + "</div>");
                    }
                }
                else if (mode == outputType.text)
                {
                    footerText.AppendLine(tm.getLabel("ftHeader"));
                    footerText.AppendLine(tm.getLabel("ftArtDecorReference") + "[http://decor.nictiz.nl/art-decor/decor-datasets--zib1bbr-?id=" + zOID +
                        "&effectiveDate=" + Convert.ToDateTime(getPublishDate()).ToString("s") + "&language=" + multilanguageText.languageCode[Settings.zibcontext.pubLanguage] + "]");
                    footerText.AppendLine(tm.getLabel("ftHeader2"));
                    footerText.AppendLine(tm.getLabel("ftInfoBase") + Settings.XGenConfig.getReleaseInfo(Settings.zibcontext.publicatie, Settings.zibcontext.PreReleaseNumber.ToString(), Settings.zibcontext.pubLanguage));
                    footerText.AppendLine(tm.getLabel("ftConditions"));
                    footerText.AppendLine(string.Format(tm.getWikiLabel("ftDate"),
                            DateTime.Now.ToString(new CultureInfo("en-GB")),
                            typeof(ZibExtraction).Assembly.GetName().Version));
                    footerText.AppendLine("Category: ZIB");
                }
                else
                    // Geen footer in XLS output
                    footerText.Append("");
                return footerText.ToString();
            }

            // ======================================================================
            // Formeren van de lijst met concepten van het informatiemodel van de ZIB
            // ======================================================================

            private void getConceptlist(EA.Package modelPackage)
            {
                conceptList.Clear();
                int indentLevel = 0;
                if (modelPackage.Elements.Count > 0)
                {
                    EA.Element rootConcept = getRootConcept(modelPackage);
                    root = getConceptDetails(rootConcept, indentLevel);
                    addConcepts(rootConcept, null, indentLevel);
                }
            }


            private void addConcepts(EA.Element e, EA.Element parent, int indentLevel)
            {
                bool includeElement = false;
                string name = e.Name;
                EA.Element child = null;
                indentLevel++;
                foreach (EA.Connector c in e.Connectors)
                {
                    int j = e.Connectors.Count;
                    //               Alleen connectors die niet naar de parent lopen
                    if (parent == null || (c.SupplierID != parent.ElementID && c.ClientID != parent.ElementID))
                    {
                        // dit element is eindpunt van de connenctor
                        if (c.SupplierID == e.ElementID)
                        {
                            child = r.GetElementByID(c.ClientID);
                            includeElement = true;
                        }
                        // dit element is beginpunt van de connector: alleen bij boundaries
                        else if (c.ClientID == e.ElementID)
                        {
                            child = r.GetElementByID(c.SupplierID);
                        // de eerste conditie is voor oude bouwstenen en gebaseerd op de aanwezigheid van het woord 'Bouwsteen' in de boundary name
                        // de tweede conditie is voor nieuwe bouwstenen met een HCIM::BoundaryType tag met waarde "HCIM"
                            if ((child.Type == "Boundary" && child.Name.Contains("Bouwsteen")) ||
                                (child.TaggedValuesEx.GetByName("HCIM::BoundaryType")?.Value == "HCIM"))
                                includeElement = true; else includeElement = false;
                        }
                        if (includeElement)
                        {
                            if (c.Type == "Aggregation")
                            {
                                if (child.Type != "Boundary") conceptList.Add(getConceptDetails(child, indentLevel));
                                addConcepts(child, e, indentLevel);
                            }
                        }
                        //                   else break;  // uitgecommentarieerd ivm afbreken lijstopbouw in Medicatievoorschrift tgv notelink op containerniveau
                    }
                }
            }

            private concept getConceptDetails(EA.Element e, int indentLevel)
            {
                concept c = new concept();
                c.maxIndent = maxIndent;
                c.name = e.Name;
                c.alias = e.Alias;
                c.id = e.ElementID;
                c.conceptCode = getTaggedValues(e, tagType.ZIBCode).ConvertAll(o => new conceptTag { Value = o.Value, Notes = o.Notes });
                c.definitionCodes = getTaggedValues(e, tagType.DefinitionCode).ConvertAll(o => new definitionTag { Value = o.Value, Notes = o.Notes });
                c.definition = e.Notes;
                c.indentLevel = indentLevel;
                c.type = getDatatype(e);
                c.stereotype = e.StereotypeEx;
                c.cardinality = getCardinality(e);
                c.constraint = getConstraints(e);
                c.valuelists = getTaggedValues(e, tagType.Valueset).ConvertAll(o => new valuesetTag { Value = o.Value, Notes = o.Notes });
                c.authorities = getTaggedValues(e, tagType.AssigningAuthority).ConvertAll(o => new authorityTag { Value = o.Value, Notes = o.Notes });
                c.zibrefs = getTaggedValues(e, tagType.ZIBReference).ConvertAll(o => new zibrefTag { Value = o.Value, Notes = o.Notes });
                c.treepos = e.TreePos;
                c.packageId = e.PackageID; ;
                if (c.stereotype == "rootconcept")
                    c.concepttype = conceptType.rootconcept;
                else if (c.stereotype == "container")
                    c.concepttype = conceptType.container;
                else if (c.stereotype.Contains("reference"))
                    c.concepttype = conceptType.reference;
                else if (string.IsNullOrWhiteSpace(c.stereotype))
                    c.concepttype = conceptType.error;
                else
                    c.concepttype = conceptType.data;
                return c;
            }



            private EA.Element getRootConcept(EA.Package modelPackage)
            {
                EA.Element e = default;
                for (short i = 0; i < modelPackage.Elements.Count; i++)
                {
                    if (modelPackage.Elements.GetAt(i).Stereotype == "rootconcept")
                    {
                        e = modelPackage.Elements.GetAt(i);
                    }
                }
                return e;
            }

            // =========================================
            // methodes per concept attribute
            // =========================================

            private string getDatatype(EA.Element e)
            {
                string dataType = "";
                foreach (EA.Connector con in e.Connectors)
                {
                    if (con.Type == "Generalization")
                    {
                        dataType = r.GetElementByID(con.SupplierID).Name;
                    }
                }
                return dataType;
            }

            private string getCardinality(EA.Element e)
            {
                string card = "";
                foreach (EA.Connector con in e.Connectors)
                {
                    // Cardinaliteit van de connector met type Aggregation waarvan het huidige element de source is. 
                    if (con.Type == "Aggregation" && con.ClientID == e.ElementID)
                    {
                        EA.Element t = r.GetElementByID(con.SupplierID);
                        if (t.Type == "Class")
                           card += con.ClientEnd.Cardinality;
                        else if (t.Type == "Boundary" && (t.TaggedValuesEx.GetByName("HCIM::BoundaryType")?.Value == "ChoiceBox"))
                            card += "(0..1)";
                    }
                }
                return card;
            }

            private List<string> getConstraints(EA.Element e)
            {
                List<string> constraintList = new List<string>();
                EA.Element constraint;

                //          Constraints die in de concepten zijn opgenomen
                foreach (EA.Constraint c in e.Constraints)
                {
                    constraintList.Add(c.Name);
                }

                //          Constraints die als losse concepten zijn opgenomen

                foreach (EA.Connector con in e.Connectors)
                {
                    if (con.Type == "NoteLink")
                    {
                        constraint = r.GetElementByID(con.ClientID);
                        if (constraint.Type == "Constraint")
                        {
                            constraintList.Add(constraint.Notes);
                        }
                    }

                }
                return constraintList;
            }

            private string getZibLifeCycleStatus()
            {
                string _Status = "";
                foreach (EA.TaggedValue tag in zp.Element.TaggedValuesEx)
                    if (tag.Name == "DCM::LifecycleStatus") _Status = tag.Value;
                return _Status;
            }


            private string getZibPublicationStatus()
            {
                string _PublicationStatus = "";
                foreach (EA.TaggedValue tag in zp.Element.TaggedValuesEx)
                    if (tag.Name == "DCM::PublicationStatus") _PublicationStatus = tag.Value;
                return _PublicationStatus;
            }
            private string getZibPublicationDate()
            {
                string _PublicationDate = "";
                foreach (EA.TaggedValue tag in zp.Element.TaggedValuesEx)
                    if (tag.Name == "DCM::PublicationDate") _PublicationDate = tag.Value;
                return _PublicationDate;
            }


            private List<EA.TaggedValue> getTaggedValues(EA.Element e, tagType qtag)
            {
                List<EA.TaggedValue> conceptTags = new List<EA.TaggedValue>();
                // lijst met lambda functions
                var qdict = new Dictionary<tagType, Func<EA.TaggedValue, bool>>
            {
                { tagType.ZIBCode, (ctag) => (ctag.Name == "DCM::DefinitionCode" && ctag.Value.Contains("NL-CM")) || ctag.Name == "DCM::ConceptId" },
                { tagType.DefinitionCode, (ctag) => (ctag.Name == "DCM::DefinitionCode" && !ctag.Value.Contains("NL-CM")) },
                { tagType.Valueset, (ctag) => (ctag.Name == "DCM::ValueSet") },
                { tagType.AssigningAuthority, (ctag) => ctag.Name == "DCM::AssigningAuthority" },
                { tagType.ZIBReference, (ctag) => ctag.Name == "DCM::ReferencedDefinitionCode" || ctag.Name == "DCM::ReferencedConceptId"},
                { tagType.Example, (ctag) => ctag.Name == "DCM::ExampleValue" },
            };
                conceptTags.Clear();
                //  zet de tags om in een c#  list. LINQ werkt niet op EA collections
                foreach (EA.TaggedValue tag in e.TaggedValuesEx)
                {
                    conceptTags.Add(tag);
                }
                // doe de LINQ queries
                var Tags = from conceptTag in conceptTags
                           where qdict[qtag](conceptTag)
                           orderby conceptTag.Value ascending
                           select conceptTag;
                return Tags.ToList();
            }

            // ==============
            // Waardelijsten
            // ==============

            public string getValueSets(EA.Package p, outputType mode)
            {
                List<string[]> vlLists = new List<string[]>();
                StringWriter valuesetText = new StringWriter();
                string vsPath = Settings.userPreferences.VS_RTFLocation;
                XDocument XValueSet = null;
                XElement Xsets = null;

                dumpValuesets(p, ref vlLists);

                // RTF files uit geheugen lezen gaat niet, dus worden ze eerst opgeslagen en dan gelezen. 9-6 Nu wel uit geheugen

                Valueset VS = new Valueset();
                VS.NewMessage += valuesetMessage;
                if (mode == outputType.xml || mode == outputType.xmlAD)
                {
                    XValueSet = new XDocument(
                        new XDeclaration("1.0", "", ""),
                        new XComment("Waardenlijsten van ZIB " + Fullname)
                    );
                    Xsets = new XElement("terminology");
                    XValueSet.Add(Xsets);
                }

                //          Sorteer de waardelijsten alfabetisch
                IEnumerable<string[]> sortAscendingQuery =
                    from vlList in vlLists
                    orderby vlList[6]
                    select vlList;

                foreach (string[] vl in sortAscendingQuery)
                {
//                    string fileName = vsPath + vl[3] + @".rtf";
                    VS.tagName = vl[6];
                    VS.tagOID = vl[7];
                    VS.tagBinding = vl[8];
                    VS.tagLanguage = vl[10];
                    //17-12-24 twee extra tags
                    VS.tagStatus = vl[11];
                    bool validParse = bool.TryParse(vl[12], out bool tmp);
                    if (validParse)
                        VS.tagIncludeOTH = tmp;
                    else
                        VS.tagIncludeOTH = null;
//                    VS.ReadContent(fileName);    van file naar string
                    VS.ReadContent(vl[3]);
                    VS.Notes = vl[9];
                    if (mode == outputType.wiki)
                        valuesetText.WriteLine(VS.Convert2Wiki());
                    else if (mode == outputType.xml || mode == outputType.xmlAD)
                    {
                        VS.EA_element_GUID = vl[2];
                        VS.EA_element_ID = vl[1];
                        VS.EA_doc_GUID = vl[5];
                        VS.EA_doc_ID = vl[4];
                        if (mode == outputType.xml)
                            VS.Convert2XML(ref Xsets);
                        else
                            VS.Convert2XML_AD(ref Xsets, getPublishDate());
                        valuesetText.WriteLine(zName + ": Waardelijst " + vl[6]);
                    }
                    else if (mode == outputType.xls)
                        valuesetText.Write(VS.Convert2XLS());
                    else
                        valuesetText.WriteLine(string.Join(", ", vl));
                }
                if (mode == outputType.xml || mode == outputType.xmlAD)
                    XValueSet.Save(Path.Combine(Settings.userPreferences.XMLLocation, zibName.fileName(zFullname, currentLanguage) + "_valuesets.xml"));
                return valuesetText.ToString();
            }

            private void dumpValuesets(EA.Package p, ref List<string[]> vsLists)
            {
                // 17-12-24
                string vsDocument, vsOID, vsBinding, vsLanguage, vsStatus, vsIncludeOTH;
                string vsPath = Settings.userPreferences.VS_RTFLocation;
                string datatype;

                foreach (EA.Element e in p.Elements)
                {
                    datatype = getDatatype(e);
                    if (datatype == "CD" || datatype == "CO" || datatype == "ANY")
                    {
                        foreach (EA.Connector c in e.Connectors)
                        {
                            EA.Element eDoc = r.GetElementByID(c.ClientID);
                            if (c.SupplierID == e.ElementID && c.Type == "Dependency" && eDoc.Stereotype == "document")
                            {
                                //                                string valuesetFullName = eDoc.Name + "-v" + zVersion;
                                //                                vsLists.Add(new string[8] { e.Name, e.ElementID.ToString(), e.ElementGUID, zibName.fileName(valuesetFullName, currentLanguage), eDoc.ElementID.ToString(), eDoc.ElementGUID, eDoc.Name, "" });
                                vsOID = eDoc.TaggedValuesEx.OfType<EA.TaggedValue>()?.Where(x => x.Name == "DCM::ValueSetId")?.FirstOrDefault()?.Value;
                                vsBinding = eDoc.TaggedValuesEx.OfType<EA.TaggedValue>()?.Where(x => x.Name == "DCM::ValueSetBinding")?.FirstOrDefault()?.Value ?? "";
                                // 17-12-24 twee extra tags
                                vsStatus = eDoc.TaggedValuesEx.OfType<EA.TaggedValue>()?.Where(x => x.Name == "DCM::ValueSetStatus")?.FirstOrDefault()?.Value ?? "";
                                vsIncludeOTH = eDoc.TaggedValuesEx.OfType<EA.TaggedValue>()?.Where(x => x.Name == "DCM::ValueSetIncludeOTH")?.FirstOrDefault()?.Value ?? "";
                                vsLanguage = eDoc.TaggedValuesEx.OfType<EA.TaggedValue>()?.Where(x => x.Name == "HCIM::ValueSetLanguage")?.FirstOrDefault()?.Value ?? "";
                                vsDocument = eDoc.GetLinkedDocument();
                                //17-12-24 2 extra waarden, van 11 naar 13
                                vsLists.Add(new string[13] { e.Name, e.ElementID.ToString(), e.ElementGUID, vsDocument, eDoc.ElementID.ToString(), eDoc.ElementGUID, eDoc.Name, vsOID, vsBinding, eDoc.Notes, vsLanguage, vsStatus, vsIncludeOTH });
// Nieuwe filenamen                                string vsFilename = eDoc.Name + @".rtf";
//                                string vsFilename = zibName.fileName(valuesetFullName, currentLanguage) + @".rtf";
//                                File.WriteAllText(vsPath + vsFilename, vsDocument);
                            }
                        }
                    }
                }
                foreach (EA.Package childPackage in p.Packages)
                    dumpValuesets(childPackage, ref vsLists);
            }

            // =====================================
            // Assigning authorities
            // =====================================

            private string createAssigningAuthorities(EA.Package p, outputType mode)
            {
                StringBuilder temp = new StringBuilder();
                IEnumerable<authorityTag> tList;

                tList = getAssigningAuthorities(p);

                if (tList != null && tList.Count() != 0)
                {
                    if (mode == outputType.wiki)
                    {
                        foreach (authorityTag t in tList)
                        {
                            temp.AppendLine("=== " + t.Value + "===");
                            temp.AppendLine("{| class=\"wikitable\" width=\"60%\"");
                            temp.AppendLine("|-style=\"background-color: #1F497D; color: white; font-weight: bold;\"");
                            temp.AppendLine("|style=\"width: 50%;\"|" + tm.getWikiLabel("caId") + "||" + tm.getWikiLabel("caSystemOID"));
                            temp.AppendLine("|-style=\"vertical-align:top;\"");
                            temp.AppendLine("|" + t.Value + "||" + t.Notes);
                            temp.AppendLine("|}");
                        }
                    }
                    else if (mode == outputType.xls)
                    {
                        temp.AppendLine("<<Sheet>>Assigning Authorities;3;3");
                        foreach (authorityTag t in tList)
                        {
                            temp.AppendLine(tm.getLabel("caName") + " " + t.Value);
                            temp.AppendLine("<<Merge>>2");
                            temp.AppendLine("<<Header>>");
                            temp.AppendLine(tm.getLabel("caId") + ";" + tm.getLabel("caSystemOID"));
                            temp.AppendLine("<<Subheader>>");
                            temp.AppendLine(t.Value + ";" + t.Notes);
                            temp.AppendLine(" ");
                        }
                        temp.AppendLine("<<ColWidth>> Auto");
                    }
                    else
                    {
                        foreach (authorityTag t in tList)
                        {
                            temp.AppendLine(t.ToString());
                        }
                        temp.AppendLine("");
                    }
                }
                return temp.ToString();
            }

            private IEnumerable<authorityTag> getAssigningAuthorities(EA.Package p)
            {
                IEnumerable<authorityTag> tList, tempList;
                tList = null;
                tempList = null;
                foreach (EA.Element e in p.Elements)
                    if (getDatatype(e) == "II")
                    {
                        tList = getTaggedValues(e, tagType.AssigningAuthority).ConvertAll(o => new authorityTag { Value = o.Value, Notes = o.Notes });
                    }
                int i = p.Packages.Count;
                foreach (EA.Package child in p.Packages)
                {
                    tempList = getAssigningAuthorities(child);
                    if (tempList != null && tempList.Count() != 0) tList = tList.Concat(tempList);
                }
                return tList;
            }


            // ================================
            // Informatiemodel diagrammen
            // ================================

            private Dictionary<int, string[]> getDiagrams(EA.Package p, textLanguage language)
            {
                Dictionary<int, string[]> diagramList = new Dictionary<int, string[]>();
                dumpDiagram(p, ref diagramList, language);
                return diagramList;
            }

            /// <summary>
            /// Geeft voor een diagram de coordinaten van de classes en artifacts (= codelijsten) tbv clickable imagemaps.
            /// Bij export van het diagram naar een png maakt EA de bovenmargin 35 pixels en de linksmargin 25 pixels. (versie 14: 43 resp 35 pixels)
            /// Als geen border getekend wordt zjn beide marges (versie 14) 10 pixels
            /// Om de werkelijke positie van een element in de png te berekenen, worden de coordinaten gecorrigeerd met het verschil tussen de werkelijke afstand
            /// van het bovenste en van het meeste linkse element tov het punt 0,0 en de hierbovengenoemde genoemde vaste marges.
            /// </summary>
            /// <param name="diagramID" Id van het diagram></param>
            /// <returns>Dictonary met index elementId en als waarde de topleft en bottomright points</returns>
            private Dictionary<int, imageMapData> getDiagramObjectsLinks(int diagramID)
            {
                Dictionary<int, imageMapData> _diagramObjects = new Dictionary<int, imageMapData>();
                EA.Element _element;
                string zibID, _zibName, link;
                EA.Diagram _diagram = r.GetDiagramByID(diagramID);
                int version;
                try
                {
                    version = int.Parse(Settings.EA_Version.Split('.')[0]);
                }
                catch
                {
                    version = 1;
                }
                int id =  _diagram.Scale;
                int minimalTop = _diagram.DiagramObjects.OfType<EA.DiagramObject>().Max(y => y.top);
                int minimalLeft = _diagram.DiagramObjects.OfType<EA.DiagramObject>().Min(y => y.left);
                int xCorrection = minimalLeft - 25 - (version > 13 ? 10 : 0);  //was 25
                int yCorrection = minimalTop + 35 + (version > 13 ? 8 : 0);   //was 35
//                int xCorrection = minimalLeft - 10;
//                int yCorrection = minimalTop + 10;
                OnNewMessage(new ErrorMessageEventArgs(ErrorType.information, "Scale: " + _diagram.Scale));

                foreach (EA.DiagramObject dO in _diagram.DiagramObjects)
                {
                    _element = r.GetElementByID(dO.ElementID);
                    if (_element.Type == "Class" || _element.Type == "Artifact")
                    {
                        link = "";
                        if (_element.StereotypeEx.Contains("reference"))
                        {
                            zibID = _element.TaggedValuesEx.OfType<EA.TaggedValue>().Where(x => x.Name == "DCM::ReferencedConceptId").First().Value;
                            _zibName = Settings.XGenConfig.getLanguageSpecificName(zibID, Settings.zibcontext.publicatie, Settings.zibcontext.PreReleaseNumber.ToString(), Settings.zibcontext.pubLanguage);
                            link = zibName.wikiLink(_zibName, Settings.zibcontext.pubLanguage);
                        }
                        bool template = Settings.zibcontext.zibPrefix.Contains("template");
                        _diagramObjects.Add(dO.ElementID, new imageMapData() { topLeft = new Point(dO.left - xCorrection, dO.top - yCorrection), bottomRight = new Point(dO.right - xCorrection, dO.bottom - yCorrection), link = _element.Type == "Class" ? (link == "" ? "#" + dO.ElementID.ToString() : link) : "#" + (template? _element.Name.TemplateNameToWiki(): _element.Name) });
                    }
                }
                return _diagramObjects;
            }


            private void dumpDiagram(EA.Package p, ref Dictionary<int, string[]> diagramList, textLanguage language)
            {
                string path = Settings.userPreferences.DiagramLocation;
                string publication = Settings.zibcontext.publicatie ?? "??";
                bool includePublication = !Settings.zibcontext.zibPrefix.Contains("template");
                foreach (EA.Diagram currentDiagram in p.Diagrams)
                {
                    string filename = currentDiagram.Name;
                    int id = currentDiagram.DiagramID;
                    //                    if (filename == "Information Model") filename = shortname;
                    if (filename == "Information Model") filename = zShortname + "-v" + zVersion; else filename += "-v" + zVersion;
                    //filename = filename + "Model" + (language == textLanguage.Multi ? "" : "(" + language.ToString() + ")") + ".png";
                    filename = filename + "Model" + (language == textLanguage.Multi ? "" : "(" + (includePublication ? publication : "") + language.ToString() + ")") + ".png";
                    diagramList.Add(id, new string[] {currentDiagram.Name, filename});
                    filename = path + filename;
                    EA.Project project = new EA.Project();
                    project = r.GetProjectInterface();
                    project.PutDiagramImageToFile(project.GUIDtoXML(currentDiagram.DiagramGUID), filename, 1);
                }

                foreach (EA.Package childPackage in p.Packages)
                {
                    dumpDiagram(childPackage, ref diagramList, language);
                }

            }

            public void getXMI(textLanguage language)
            {
                string fileName = Path.Combine(Settings.userPreferences.XMLLocation, zFullname + (language == textLanguage.Multi ? "" : "(" + language.ToString() + ")") + ".xmi");
                string zibGUID = r.GetProjectInterface().GUIDtoXML(zp.PackageGUID);
                //  Definitie: ExportPackageXMI(string PackageGUID, EnumXMIType XMIType, int DiagramXML, int DiagramImage, int FormatXML, int UseDTD, string FileName);
                r.GetProjectInterface().ExportPackageXMI(zibGUID, EA.EnumXMIType.xmiEA11, 2, -1, 1, 0, fileName);
                _ = PostProcessXMI(fileName);
            }


            /// <summary>
            /// Zibs die waardenlijsten bevatten die deprecated zijn, hebben een binding 'Deprecated'. Hoewel dat in de zibs zo afgesproken is,
            /// is de waarde 'Deprecated' Fhir niet toegestaan. In de waardenlijst XML bestanden wordt dit daarom aangepast, maar in de XMI moet deze 
            /// waarde voor de export tbv Fhir ook verwijderd te worden.
            /// </summary>
            /// <param name="fileName">Naam van de XMI file</param>
            /// <returns>succes</returns>
            public bool PostProcessXMI(string fileName)
            {
                bool result = false;
                try
                {
                    XDocument xmi = XDocument.Load(fileName);
                    XNamespace elementNamespace = xmi.Root.GetNamespaceOfPrefix("UML").NamespaceName;
                    var deprecatedBindingTags = xmi.Descendants(elementNamespace + "TaggedValue")?.Where(x => x.Attribute("tag").Value == "DCM::ValueSetBinding" && x.Attribute("value")?.Value == "Deprecated");
                    if (deprecatedBindingTags.Any())
                    {
                        foreach (XElement tag in deprecatedBindingTags)
                            tag.Attribute("value").SetValue("");
                        OnNewMessage(new ErrorMessageEventArgs(ErrorType.information, "Valuesetbinding tag 'Deprecated' aangepast"));
                        xmi.Save(fileName);
                        result = true;
                    }
                }
                catch (IOException)
                {
                    OnNewMessage(new ErrorMessageEventArgs(ErrorType.error, "Openen XMI bestand voor controle valuesetbinding is fout gegaan. Geen aanpassingen uitgevoerd."));
                    return false;
                }
                return result;
            }


            public void printToRTF(textLanguage language)
            {
                UpdateProjectConstants(Status, PublicationStatus, PublicationDate);
                OnNewMessage(new ErrorMessageEventArgs(ErrorType.information, "Projectconstanten aangepast"));
                string fileName = Path.Combine(Settings.userPreferences.ZIB_RTFLocation, zibName.fileName2(zFullname, currentLanguage) + (Settings.PDFFormat ? ".pdf" : ".docx"));
                //	Let op! In tegenstelling tot wat de documentatie zegt, moet de GUID hier niet naar XML formaat omgezet worden. Vreemd!
                string zibPackageGUID = zp.PackageGUID;
                // was:  r.GetProjectInterface().RunReport(zibPackageGUID, Settings.zibcontext.zibTemplate, fileName);
                EA.Project project = new EA.Project();
                project = r.GetProjectInterface();
                project.RunReport(zibPackageGUID, Settings.zibcontext.zibTemplate, fileName);
            }

            public void UpdateProjectConstants(string zibStatus, string zibPublicationStatus, string zibPublicationDate)
            {
                string strSQLresult = "";
                Dictionary<string, string> projectConstantsDict = new Dictionary<string, string>();
                try
                {
                    strSQLresult = r.SQLQuery("SELECT Template FROM t_rtf Where Type = 'ProjectOpts'");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Fout " + ex.ToString());
                }

                XDocument doc = XDocument.Parse(strSQLresult);
                if (doc.Descendants("Template").Count() != 0)
                {
                    string[] projectConstant = (doc.Descendants("Template").First().Value).TrimEnd(';').Split(';');
                    foreach (string s in projectConstant)
                    {
                        string[] kvp = s.Split('=');
                        if (kvp.Count() == 2)
                            projectConstantsDict.Add(kvp[0], kvp[1]);
                        else
                            MessageBox.Show("Projectconstant error: " + s);
                    }
                    projectConstantsDict["DCM_LifecycleStatus"] = zibStatus;
                    projectConstantsDict["DCM_Name"] = zFullname;
                    projectConstantsDict["DCM_PublicationStatus"] = zibPublicationStatus;
                    projectConstantsDict["DCM_PublicationDate"] = zibPublicationDate;
                    projectConstantsDict["HCIM_Release"] = Settings.zibcontext.publicatie ?? "";

                    string SQLString = "";
                    foreach (KeyValuePair<string, string> pC in projectConstantsDict)
                        SQLString += pC.Key + "=" + pC.Value + ";";

                    string SQLUpdate = "Update t_rtf Set Template = '" + SQLString + "' where Type = 'ProjectOpts'";

 //                   string SQLInsert = "Insert into t_rtf(Type, Template) values('ProjectOpts', '" + SQLString + "')";


                    //	Repository.Execute is een niet gedocumenteerde functie, maar het werkt. 
                    r.Execute(SQLUpdate);
                }
                else
                {
                    projectConstantsDict.Add("DCM_LifecycleStatus", zibStatus);
                    projectConstantsDict.Add("DCM_Name", zFullname);
                    projectConstantsDict.Add("DCM_PublicationStatus", zibPublicationStatus);
                    projectConstantsDict.Add("DCM_PublicationDate", zibPublicationDate);
                    projectConstantsDict.Add("HCIM_Release", Settings.zibcontext.publicatie ?? "");

                    string SQLString = "";
                    foreach (KeyValuePair<string, string> pC in projectConstantsDict)
                        SQLString += pC.Key + "=" + pC.Value + ";";

                    string SQLInsert = "Insert into t_rtf(Type, Template) values('ProjectOpts', '" + SQLString + "')";
                    r.Execute(SQLInsert);
                }
            }

            public bool getSingleLanguageVersion() // succes nog teruggeven
            {
                bool succes = false;
                string definitionSingle;
                //
                // Release 2015 en 2016 hebben niet altijd een alias
                if (!Settings.oldFormat) succes = swapNameAndAlias(zp.Element);
                //

                foreach (EA.Package childPackage in zp.Packages)
                {
                    bool deselectedPackage;
                    deselectedPackage = (childPackage.Name == "Revision History");
                    deselectedPackage = deselectedPackage || (childPackage.Name == "Mindmap");
                    deselectedPackage = deselectedPackage || (childPackage.Name == "Information Model");
                    deselectedPackage = deselectedPackage || (childPackage.Name == "Example Instances");
                    deselectedPackage = deselectedPackage || (childPackage.Name == "Example of the Instrument");
                    deselectedPackage = deselectedPackage || (childPackage.Name == "References");
                    if (!deselectedPackage)
                    {
                        multilanguageText definitionMulti = (multilanguageText)childPackage.Notes;
                        try
                        {
                            definitionSingle = definitionMulti.toSingle(Settings.zibcontext.pubLanguage);
                            childPackage.Notes = definitionSingle;
                            childPackage.Update();
                            succes = true;
                        }
                        catch (InvalidCastException e)
                        {
                            reportError(ErrorType.error, e.Message);
                            succes = false;
                        }
                        succes = getElementTaggedValues(childPackage.Element);
                    }
                    int i = childPackage.Elements.Count;
                    int j = childPackage.Element.Elements.Count;

                    succes = GetPackageElements(childPackage);
                }
                if (succes) setZibPublishLanguage(Settings.zibcontext.pubLanguage);
                zp.Name = this.Fullname;
                zp.Update();
                return succes;
            }

            bool GetPackageElements(EA.Package theChildPackage)
            {
                EA.Element theElement;
                bool succes = true;
                if (theChildPackage.Elements.Count > 0)
                {
                    succes = false;
                    for (short i = 0; i < theChildPackage.Elements.Count; i++)
                    {
                        theElement = theChildPackage.Elements.GetAt(i);
                        succes = GetSingleLanguageElement(theElement);
                    }
                }
                //                	  recurse	
                foreach (EA.Package grandchildPackage in theChildPackage.Packages)
                {
                    succes &= swapNameAndAlias(grandchildPackage.Element);
                    succes &= GetPackageElements(grandchildPackage);
                }
                return succes;
            }

            private bool GetElementElements(EA.Element theChildElement)
            {
                bool succes = true;
                EA.Element theElement = null;
                if (theChildElement.Elements.Count > 0)
                {
                    succes = false;
                    for (short i = 0; i < theChildElement.Elements.Count; i++)
                    {
                        theElement = theChildElement.Elements.GetAt(i);
                        succes = GetSingleLanguageElement(theElement);
                    }
                }
                foreach (EA.Element grandchildElement in theChildElement.Elements)
                {
                    succes &= GetElementElements(grandchildElement);
                }
                return succes;
            }

            bool GetSingleLanguageElement(EA.Element theElement)
            {
                bool selectedObjectType;
                bool succes = false;
                string elementName;
                string definitionSingle;

                elementName = theElement.Name;

                selectedObjectType = (theElement.Type == "Class");
                selectedObjectType = selectedObjectType || (theElement.Type == "Constraint");
                selectedObjectType = selectedObjectType || (theElement.Type == "Note");
                selectedObjectType = selectedObjectType || (theElement.Type == "Artifact");
                // 30/8/2022 Boundary toegevoegd
                selectedObjectType = selectedObjectType || (theElement.Type == "Boundary");

                if (selectedObjectType)
                {
                    multilanguageText definitionMulti = (multilanguageText)theElement.Notes;
                    try
                    {
                        definitionSingle = definitionMulti.toSingle(Settings.zibcontext.pubLanguage);
                        theElement.Notes = definitionSingle;
                        theElement.Update();
                        succes = true;
                    }
                    catch (InvalidCastException e)
                    {
                        reportError(ErrorType.error, e.Message);
                        succes = false;
                    }
                    succes = getElementTaggedValues(theElement);
                    // Release 2015 en 2016 hebben niet altijd een alias
                    if (!Settings.oldFormat) succes = swapNameAndAlias(theElement);
                }

                if (theElement.ConstraintsEx.Count > 0)
                {
                    // Dit zijn de constraints, die in het class diagramobject komen. Ze worden eigenlijk alleen gebruikt om het waardebereik aan te geven
                    string constraintTextSingle;

                    // het waardebereik standaard teksten
                    Dictionary<textLanguage, string> rangeText = new Dictionary<textLanguage, string>()
                    {
                        { textLanguage.NL, "Waardebereik:"},
                        { textLanguage.EN, "Range:" }
                    };

                    EA.Constraint _constraint;
                    var constraints = theElement.Constraints.OfType<EA.Constraint>().Select((x, y) => new { y, data = new string[] { x.Name, x.Notes } }).ToDictionary(z => z.y, z => z.data);

                    // stel de nieuwe lijst met constraints vast met vertaalde waardelijst tekst en notes singleLanguage tekst
                    foreach (KeyValuePair<int, string[]> c in constraints)
                    {
                        if (Settings.zibcontext.pubLanguage != textLanguage.NL) 
                        {
                            c.Value[0] = c.Value[0].Replace(rangeText[textLanguage.NL], rangeText[Settings.zibcontext.pubLanguage]);                      
                        }

                        multilanguageText constraintTextMulti = (multilanguageText)c.Value[1];
                        try
                        {
                            constraintTextSingle = constraintTextMulti.toSingle(Settings.zibcontext.pubLanguage);
                            c.Value[1] = constraintTextSingle;
                            succes = true;
                        }
                        catch (InvalidCastException e)
                        {
                            reportError(ErrorType.error, e.Message);
                            c.Value[1] = e.Message;
                            succes = false;
                        }
                    }

                    //  Verwijder de oude constraints, want de feitelijke info zit in de Name, en die kan niet geupdate worden
                    foreach (var c in theElement.Constraints.OfType<EA.Constraint>().Select((y, i) => i))
                    {
                        theElement.Constraints.Delete((short)c);
                    }
                    theElement.Constraints.Refresh();

                    // Voeg de nieuwe constraints toe
                    foreach (KeyValuePair<int, string[]> c in constraints)
                    {
                        _constraint = theElement.Constraints.AddNew(c.Value[0], "Invariant");
                        _constraint.Name = c.Value[0];
                        _constraint.Notes = c.Value[1];
                        _constraint.Update();
                    }
                    theElement.Constraints.Refresh();
                }

                //   Hier selfgenerated definition codes vertalen.
                var conceptId = theElement.TaggedValuesEx.OfType<EA.TaggedValue>().Where(x => x.Name == "DCM::ConceptId").FirstOrDefault()?.Value;
                var definitionCodeTag = theElement.TaggedValuesEx.OfType<EA.TaggedValue>().Where(x => x.Name == "DCM::DefinitionCode").FirstOrDefault();
                if (definitionCodeTag != null)
                {
                    string[] codeParts = definitionCodePart(definitionCodeTag.Value);
                    // splits en test of de laatste afgeleid is
                    if (IsAutoCode(conceptId, codeParts[1]))
                    {
                        var rootName = getRootConcept(zp.Packages.GetByName("Information Model")).Name;
                        // componeer nieuwe naam, rootconcept naam voor nodig
                        definitionCodeTag.Value = codeParts[0] + ": " + codeParts[1] + " " + rootName + " " + theElement.Name;
                        definitionCodeTag.Update();
                    }
                }

                //                		en  de subelementen		
                if (theElement.Elements.Count > 0)
                    GetElementElements(theElement);
                return succes;
            }


            private bool IsAutoCode(string conceptId, string definitionCode)
            {
                string[] parts;
                try
                {
                    parts = conceptId.Split(':');
                    parts = parts[1].Split('.');
                }
                catch
                {
                    return false;
                }
                return definitionCode == parts[0] + parts[1].PadLeft(3, '0') + parts[2].PadLeft(3, '0');
            }

            private string[] definitionCodePart(string definitionCodeString)
            {
                string[] codeParts = new string[3];
                int colon = definitionCodeString.IndexOf(":");
                codeParts[0] = colon == -1 ? definitionCodeString : definitionCodeString.Substring(0, colon);
                if (colon != -1 && definitionCodeString.Length > colon + 1)
                {
                    string remainder = definitionCodeString.Substring(colon + 1).Trim();
                    int space = remainder.IndexOf(" ");
                    codeParts[1] = remainder.Substring(0, space == -1 ? remainder.Length : space);
                    codeParts[2] = remainder.Substring(space + 1);
                }
                return codeParts;
            }

            private bool getElementTaggedValues(EA.Element theElement)
            {
                string tagNote_S;
                bool succes = false;

                if (theElement.TaggedValuesEx.Count > 0)
                {
                    foreach (EA.TaggedValue tag in theElement.TaggedValuesEx)
                    {
                        multilanguageText tagNote = (multilanguageText)tag.Notes;
                        try
                        {
                            tagNote_S = tagNote.toSingle(Settings.zibcontext.pubLanguage);
                            tag.Notes = tagNote_S;
                            tag.Update();
                            succes = true;
                        }
                        catch (InvalidCastException e)
                        {
                            reportError(ErrorType.error, e.Message);
                            succes = false;
                        }
                    }
                }
                else
                    succes = true;
                return succes;
            }

            public bool resizeDiagramTextboxes()
            {
                /* Na het eentalig maken van de zib zijn de textboxes van Notes en Constraint te groot
                 * Deze methode brengt de onderkant van de box terug tot de maat die op grond van de eentalige tekst nodig is.
                 * Hiervoor worden alle diagrammen in de sectie Information Model aangepast.
                 * EA zorgt ervoor dat de box niet kleiner kan zijn dan de tekst toelaat.
                 */

                bool succes = false;
                EA.Package informationModel = zp.Packages.GetByName("Information Model");
                succes = updateDiagram(informationModel);
                return succes;
            }

            private bool updateDiagram(EA.Package thePackage)
            {
                /* Deze methode doet het feitelijke werk voor de methode resizeDiagramTextboxes.
                 * Tevens wordt gecontroleerd of het beginpunt van de connector nog binnen de contouren van de textbox ligt.
                 * Zo niet wordt dit aangepast. 
                 * Verder wordt de methode recursief aangeroepen voor alle onderliggende packages
                 * hij retourneert true als alle updates van de textboxen gelukt zijn.
                 */
                bool succes = true;
                // De hoogte bepaling van de constraints en notes gaat er van uit dat EA deze schrijft in Calibri 8pt.
                Font textFont = new Font("Calibri", 8.0f);
                foreach (EA.Diagram theDiagram in thePackage.Diagrams)
                {
                    foreach (EA.DiagramObject diagramElement in theDiagram.DiagramObjects)
                    {
                        EA.Element Element = r.GetElementByID(diagramElement.ElementID);
                        if (Element.Type == "Constraint" || Element.Type == "Note")
                        {
                            TextFormatFlags flags = TextFormatFlags.WordBreak;
                            Size textSize = TextRenderer.MeasureText(Element.Notes, textFont, new Size(diagramElement.right - diagramElement.left, textFont.Height + 6), flags);
                            //Debug.WriteLine("\r\n****  " + Element.Type);
                            //Debug.WriteLine("TextSIZE: width - height |  " + textSize.Width + "," + textSize.Height);
                            //Debug.WriteLine("Box(org): top(x,y) - bottom(x,y) |  " + diagramElement.left + "," + diagramElement.top + " - " + diagramElement.right + "," + diagramElement.bottom + " Breedte, hoogte: " + (diagramElement.right - diagramElement.left) + ", " + (diagramElement.bottom - diagramElement.top).ToString());
                            // EA rendert notes anders dan constraints, 16/12/18 padding toegevoegd
                            //diagramElement.bottom = diagramElement.top - (textSize.Height + 6) - (Element.Type == "Note" ? textFont.Height : 0);
                            //18-01-22: Hoogte bepaling versimpeld, dat blijk nu wel te werken. De box wordt opzettelijk te klein gemaakt en daarna door EA automatisch aangepast.
                            diagramElement.bottom = diagramElement.top - textFont.Height;
                            if (diagramElement.right - diagramElement.left < textSize.Width) diagramElement.right = diagramElement.left + textSize.Width;
                            succes &= diagramElement.Update();
                            succes &= theDiagram.Update();
                            r.SaveDiagram(theDiagram.DiagramID);


                            // Check of het beginpunt van de connector niet buiten de textbox komt en corrigeer dit eventueel.
                            Debug.WriteLine("Box(new): top(x,y) - bottom(x,y) |  " + diagramElement.left + "," + diagramElement.top + " - " + diagramElement.right + "," + diagramElement.bottom + " Breedte, hoogte: " + (diagramElement.right - diagramElement.left) + ", " + (diagramElement.bottom - diagramElement.top).ToString());
                            var conns = Element.Connectors.OfType<EA.Connector>().Where(x => x.ClientID == Element.ElementID);
                            foreach (EA.Connector conn in conns)
                            {
                                //Debug.WriteLine("Element: "+ r.GetElementByID(conn.SupplierID).Name);
                                EA.DiagramLink link = theDiagram.DiagramLinks.OfType<EA.DiagramLink>().Where(x => x.ConnectorID == conn.ConnectorID).FirstOrDefault();
                                if (link != null)
                                {
                                    string[] geometrySplit = link.Geometry.Split(';');
                                    var sx = geometrySplit.Where(x => x.Contains("SX")).FirstOrDefault()?.Split('=')[1] ?? "0";
                                    var sy = geometrySplit.Where(x => x.Contains("SY")).FirstOrDefault()?.Split('=')[1] ?? "0";
                                    var edge = geometrySplit.Where(x => x.Contains("EDGE")).FirstOrDefault()?.Split('=')[1] ?? "";
 /*                                   string side = "";
                                    switch (edge)
                                    {
                                        case "1":
                                            side = "Top";
                                            break;
                                        case "3":
                                            side = "Bottom";
                                            break;
                                        case "4":
                                            side = "Left";
                                            break;
                                        case "2":
                                            side = "Right";
                                            break;
                                        default:
                                            side = "Unknown";
                                            break;
                                    } */
                                    //Debug.WriteLine("Offset: x,y " + sx + ", " + sy + " zijde: " + side + "  x,y: " + ((diagramElement.right - diagramElement.left) / 2 + int.Parse(sx)) + "," + ((diagramElement.bottom - diagramElement.top) / 2 + int.Parse(sy)));
                                    bool doUpdate = false;
                                    int index = -1;
                                    if (Math.Abs(int.Parse(sx)) > Math.Abs((diagramElement.right - diagramElement.left) / 2))
                                    {
                                        index = geometrySplit.Select((x, y) => new { geo = x, count = y }).Where(z => z.geo.Contains("SX")).FirstOrDefault().count;
                                        geometrySplit[index] = "SX=0";
                                        doUpdate = true;
                                    }
                                    if (Math.Abs(int.Parse(sy)) > Math.Abs((diagramElement.bottom - diagramElement.top) / 2))
                                    {
                                        index = geometrySplit.Select((x, y) => new { geo = x, count = y }).Where(z => z.geo.Contains("SY")).FirstOrDefault().count;
                                        geometrySplit[index] = "SY=0";
                                        doUpdate = true;
                                    }
                                    if (doUpdate)
                                    {
                                        Debug.WriteLine("Bouwsteen: " + r.GetPackageByID(thePackage.ParentID).Name);
                                        EA.Element supplier = r.GetElementByID(conn.SupplierID);
                                        Debug.WriteLine("Element: " + (supplier.Name == "" ? supplier.Type : supplier.Name));
                                        Debug.WriteLine("Fout geconstateerd: " + link.Geometry);
                                        link.Geometry = string.Join(";", geometrySplit);
                                        Debug.WriteLine("Nieuwe geometrie: " + link.Geometry);
                                        link.Update();
                                        theDiagram.DiagramLinks.Refresh();
                                    }
                                }
                            }
                        }
                    }
                    succes &= theDiagram.Update();
                    r.SaveDiagram(theDiagram.DiagramID);
                }
                foreach (EA.Package childPackage in thePackage.Packages)
                {
                    succes &= updateDiagram(childPackage);
                }
                textFont.Dispose();
                return succes;
            }
            bool swapNameAndAlias(EA.Element theElement)
            {

                /*  Verwissel van het element de naam en het alias in de publicatietaal. 
                 *  Voeg de Nederlandse naam toe aan de aliasen al NL:naam
                 *  Als het element het hoofdpackage is van de bouwsteen, wijzig dan ook alle zibnaam voorkomens
                 */

                bool succes = false;
                bool elementZib; // Is het element van het zib package?
                int theParentId;
                EA.Element theParent;

                string nlName;

                if (Settings.zibcontext.pubLanguage == textLanguage.NL)
                    // De bouwsteen is in het Nederlands geschreven, dus val er niets te verwisselen
                    return true;

                if (theElement.Name.StartsWith(Settings.zibcontext.zibPrefix)) elementZib = true; else elementZib = false;
                if (elementZib)
                    nlName = this.Name;
                else
                    nlName = theElement.Name;

                string allAliases = theElement.Alias;
                if (allAliases.Length == 0 && nlName.Length != 0)
                {
                    reportError(ErrorType.error, "Geen alias voor " + nlName);
                    return true; // Geen reden om te stoppen
                }
                int i = allAliases.IndexOf(Settings.zibcontext.pubLanguage.ToString());
                if (i != -1)
                {
                    int j = allAliases.IndexOf(",") == -1 ? allAliases.Length - i - 3 : allAliases.IndexOf(",") - i - 4;
                    theElement.Name = allAliases.Substring(i + 3, j).Trim() + (elementZib ? "-v" + this.Version : "");
                    theElement.Alias = allAliases.Replace(allAliases.Substring(i, j + 3), "");
                    theElement.Alias = "NL: " + nlName + (theElement.Alias.Length == 0 ? "" : " ," + theElement.Alias);
                    succes = theElement.Update();
                    if (theElement.Type == "Artifact" && theElement.Stereotype == "document")
                    {    // hernoem de valuesettag
                        theParentId = theElement.Connectors.GetAt(0).SupplierID;
                        theParent = r.GetElementByID(theParentId);
                        foreach(EA.TaggedValue tag in theParent.TaggedValuesEx)
                            if(tag.Name == "DCM::ValueSet" && tag.Value == nlName)
                            {
                                tag.Value = theElement.Name;
                                succes = tag.Update();
                                theParent.TaggedValuesEx.Refresh();
                            }
                    }

                    if (succes && elementZib)
                    {
                        this.Fullname = theElement.Name;
                        this.Name = zibName.Name(this.Fullname);
                        this.Shortname = zibName.shortName(this.Fullname);
                    }
                }
                else
                {
                    if (nlName.Length != 0)
                        reportError(ErrorType.error, "Geen " + Settings.zibcontext.pubLanguage.ToString() + " alias voor " + nlName);
                    succes = true; // Geen reden om te stoppen
                }
                return succes;
            }

            public int registerZib(bool forceReg)
            {
                /*
                 * Maakt een nieuwe entry aan voor een nieuwe, nog niet geregistreerde bouwsteen. 
                 * Als forceReg true is wordt voor een al geregistreerde bouwsteen de gegevens geupdate. Alle waarden moeten meegegeven worden.
                 * Als forceReg false is worden al geregistreerde bouwsteen entries niet veranderd.
                 * Retourneert integer  die aangeeft of de registratie succesvol was en de file dus opnieuw opgeslagen moet worden.
                 * De waarden zijn: 99 = fouten bij de check in het ZibId register; 0 = mislukt; 1= gelukt, 2 = al geregistreerd, 3= geforceerd gewijzigd
                 */

                int registered = 0;
 //               int ref_registered = 0;
                Dictionary<textLanguage, string> zibNames = new Dictionary<textLanguage, string>();
                string zName = "";

                foreach (textLanguage lang in Enum.GetValues(typeof(textLanguage)))
                {
                    if (lang != textLanguage.Multi)
                    {
                        if (lang == textLanguage.NL)
                            zName = Shortname;
                        else
                            zName = zibName.shortName(getNameFromAlias(zp.Element, lang));
                        zibNames.Add(lang, zName);
                    }
                }

                // 02-02-2024: Test of het OID wel een geregistreerd Id is, zo niet return -1;
                string xmlDocName = Path.Combine(Settings.application.CodeSystemCodesLocation, Settings.application.ZibIdentifiersFilename);
                XZibIdDoc xDoc = new XZibIdDoc(xmlDocName, Settings.application.ZibIdBase);
                if (xDoc.IsValidZibId(ZibOID) !=  0)
                    return 99;

                registered = Settings.XGenConfig.registerZib(ZibOID, zVersion, zibNames, zPrefix, Settings.forceZibRegistration);

                //                if (registered > 0) ref_registered = registerReferences(Settings.forceZibRegistration);

                // Geen idee waarom de volgend regel er in zou moeten zitten. 13-10-2019
                // Settings.XGenConfig.UsesZibs(ZibOID, Settings.zibcontext.publicatie, Settings.zibcontext.pubLanguage);

                return registered; // + (ref_registered << 8);
            }

 /*           public int registerRefs(bool forceReg)
            {
                return 0;
            }
*/
            string getNameFromAlias(EA.Element theElement, textLanguage lang)
            {
                string zName = "";
                string allAliases = theElement.Alias;
                int i = allAliases.IndexOf(Enum.GetName(typeof(textLanguage), lang));
                if (i != -1)
                {
                    int j = allAliases.IndexOf(",") == -1 ? allAliases.Length - i - 3 : allAliases.IndexOf(",") - i - 4;
                    zName = allAliases.Substring(i + 3, j).Trim();

                }
                return zName;
            }

            /// <summary>
            /// </summary>
            /// <param name="forceReg">Als forceReg true is worden de al geregistreerde referenties geupdate.
            /// Als forceReg false is worden al geregistreerde referentie entries niet veranderd.</param>
            /// <returns>Integer die aangeeft of de registratie succesvol was en de file dus opnieuw opgeslagen moet worden.
            /// Hierbij gelden de volgende waarden: 0 = mislukt, 1= nieuwe registratie, 2= al geregistreerd, 3= opnieuw geregistreerd (geforceerd).
            /// Daarnaast zijn de foutmeldingen 11 = release onbekend, 12 = zib niet in release, 13 = referenced zib niet in release, 4 zib heeft geen referenties.</returns>
            public int registerReferences(bool forceReg)
            {
                List<string> references = new List<string>();
                int registered = 0;
                EA.Package model = zp.Packages.GetByName("Information Model");
                if (model!= null)
                {
                    references.AddRange(getReferences(model));
                }

                registered= Settings.XGenConfig.registerReferences(ZibOID, Settings.zibcontext.publicatie, Settings.zibcontext.PreReleaseNumber.ToString(), references, forceReg);


                return registered;
            }

            private List<string> getReferences(EA.Package parent)
            {
                const string SQL = "SELECT DISTINCT op1.Value FROM (t_object AS o2 INNER JOIN t_objectproperties AS op1 ON op1.Object_ID = o2.Object_ID)"
                            + " WHERE op1.Property = 'DCM::ReferencedConceptId' AND o2.Package_ID = ";
                List<string> references = new List<string>();
                string queryResult = "";
                try
                {
                    queryResult = r.SQLQuery(SQL + parent.PackageID.ToString());
                }
                catch (Exception ex)
                {
                    reportError(ErrorType.error, ex.ToString());
                    return null;
                }

                XDocument doc = XDocument.Parse(queryResult);
                foreach (XElement t in doc.Descendants("Value"))
                {
                    references.Add(t.Value.ToString());
                }

                foreach (EA.Package child in parent.Packages)
                    references.AddRange(getReferences(child));
                return references;
            }

            private void reportError(ErrorType type, string s)
            {
                if (!string.IsNullOrEmpty(s)) OnNewMessage(new ErrorMessageEventArgs(type, s));
            }

            private void valuesetMessage(object sender, ErrorMessageEventArgs e)
            {
                // geef de melding door aan het hoofdprogramma
                reportError(e.MessageType, e.MessageText);
            }
            private void zibexampleMessage(object sender, ErrorMessageEventArgs e)
            {
                // geef de melding door aan het hoofdprogramma
                reportError(e.MessageType, e.MessageText);
            }


            private string getTranscludeFilename(textLanguage zibLanguage)
            {
                return "Versions-" + zOID + "(" + zibLanguage + ")";
            }

        }

        public enum tagType
        {
            ZIBCode,
            DefinitionCode,
            Valueset,
            AssigningAuthority,
            ZIBReference,
            Example
        }
        public enum outputType
        {
            text,
            wiki,
            xml,
            xmlAD,
            xls
        }

        public enum conceptType
        {
            rootconcept,
            container,
            reference,
            data,
            error
        }

        public enum sectionType
        {
            Header,
            ZIBtags,
            Revision_History,
            Concept,
            Purpose,
            Evidence_Base,
            Patient_Population,
            Information_Model,
            Example_Instances,
            Instructions,
            Issues,
            References,
            Traceability_to_other_Standards,
            Assigning_Authorities,
            Valuesets,
            Disclaimer,
            Terms_of_Use,
            Copyrights,
            Footer
        }


        public class conceptTypeData

        {
            public string rowColor;
            public string image;
            public int size;
        }

        // ==================
        // Valuetag interface
        // ==================

        public interface ITag
        {

            string Value
            {
                get;
                set;
            }

            string Notes
            {
                get;
                set;
            }

            string ToWiki();

            string ToXLS();

        }

        public class zibTag : ITag
        {
            protected string tagValue;
            protected string tagNotes;
            protected string image;

            public string Value
            {
                get { return tagValue; }
                set { tagValue = value; }
            }

            public string Notes
            {
                get { return tagNotes; }
                set { tagNotes = value; }
            }

            public virtual string ToWiki()
            {
                return "|" + tagValue + "|" + tagNotes;
            }

            public override string ToString()
            {
                return string.IsNullOrWhiteSpace(tagNotes) ? tagValue : tagValue + ":" + tagNotes;
            }

            public virtual string ToXLS()
            {
                return string.IsNullOrWhiteSpace(tagNotes) ? tagValue : tagValue + ":" + tagNotes;
            }

        }


        public class conceptTag : zibTag
        {
            public override string ToWiki()
            {
                return "|" + tagValue;
            }
            public override string ToXLS()
            {
                return tagValue;
            }
        }

        public class definitionTag : conceptTag
        {
            int translateFromYear;
            public override string ToWiki()
            {
                //Hover voor codesystem toevoegen
                string wikiTag = "|" + tagValue;
                int index = tagValue.IndexOf(":");
                if (index != -1)
                {
                    CodeNames cn = new CodeNames();
                    cn.languageRefsetId_EN = Settings.application.Snomed_languageRefsetId_EN;
                    cn.languageRefsetId_NL = Settings.application.Snomed_languageRefsetId_NL;
                    cn.terminologyLink = Settings.application.TerminologyServiceLink;
                    if (int.TryParse(Settings.application.CodeTranslateFromYear, out translateFromYear)) cn.translateFromYear = translateFromYear;
                    cn.LoadZibCodeSystems(Path.Combine(Settings.application.CodeSystemCodesLocation, Settings.application.CodeSystemCodesFilename));

                    string codeSystem = tagValue.Substring(0, index).Trim();
                    int codeIndex = tagValue.Substring(index + 1).IndexOf(" ", 2); // skip evt trailing spaces
                    string code = tagValue.Substring(index + 1, codeIndex + 1).Trim();
                    string codeName = tagValue.Substring(index + codeIndex + 1).Trim();

                    //wikiTag = "|<span Title=\"Codesystem: " + codeSystem + "\">" + tagValue.Substring(index + 1) + "</span>";
                    //wikiTag = "|<span Title=\"Codesystem: " + codeSystem + "\">" + code.toCodeLink(codeSystem) + " " + codeName + "</span>";
                    wikiTag = "|" + cn.translate4Wiki(codeName, code, codeSystem, Settings.zibcontext.publicatie, Settings.zibcontext.pubLanguage.ToString(), true);
                }
                return wikiTag;
            }
            public override string ToXLS()
            {
                string xlsTag =  ""; 
                int index = tagValue.IndexOf(":");
                if (index != -1)
                {
                    CodeNames cn = new CodeNames();
                    cn.languageRefsetId_EN = Settings.application.Snomed_languageRefsetId_EN;
                    cn.languageRefsetId_NL = Settings.application.Snomed_languageRefsetId_NL;
                    cn.terminologyLink = Settings.application.TerminologyServiceLink;
                    if (int.TryParse(Settings.application.CodeTranslateFromYear, out translateFromYear)) cn.translateFromYear = translateFromYear;
                    cn.LoadZibCodeSystems(Path.Combine(Settings.application.CodeSystemCodesLocation, Settings.application.CodeSystemCodesFilename));

                    string codeSystem = tagValue.Substring(0, index).Trim();
                    int codeIndex = tagValue.Substring(index + 1).IndexOf(" ", 2); // skip evt trailing spaces
                    string code = tagValue.Substring(index + 1, codeIndex + 1).Trim();
                    string codeName = tagValue.Substring(index + codeIndex + 1).Trim();

                    xlsTag = cn.translate4XLS(codeName, code, codeSystem, Settings.zibcontext.publicatie, Settings.zibcontext.pubLanguage.ToString(), true);
                }
                return xlsTag;
            }

        }

        public class valuesetTag : zibTag
        {
            public override string ToWiki()
            {
                string _tagValue;

                image = "List2";
                _tagValue = Settings.zibcontext.zibPrefix.Contains("template") ? tagValue.TemplateNameToWiki() : tagValue;

                return "|[[Bestand: " + image + ".png | link=#" + _tagValue + "]]||[[#" + _tagValue + "|" + _tagValue + "]]";
            }

            public override string ToXLS()
            {
                return "[Hyperlink:" + ((tagValue.Length > 31) ? tagValue.Substring(0, 31) : tagValue) + "!A1]" + tagValue;
            }
        }

        public class authorityTag : zibTag
        {
            public override string ToWiki()
            {
                image = "AA2";
                return "|[[Bestand: " + image + ".png | link=#" + tagValue + "]]||[[#" + tagValue + "|" + tagValue + "]]";
            }

            public override string ToXLS()
            {
                return "[Hyperlink:'Assigning Authorities'!A1]" + tagValue;
                //                return tagValue;
            }
        }

        public class zibrefTag : zibTag
        {

            public override string ToWiki()
            {
                image = "block"; // "zib";
                string _zibName = Settings.XGenConfig.getLanguageSpecificName(tagValue, Settings.zibcontext.publicatie, Settings.zibcontext.PreReleaseNumber.ToString(), Settings.zibcontext.pubLanguage);
//                int s = tagNotes.LastIndexOf(" ") + 1;
//                string zref = tagNotes.Substring(s, tagNotes.Length - s - (tagNotes.LastIndexOf(".") == -1 ? 0 : 1));
// 16-2                return "|[[Bestand: " + image + ".png]]||[[" + zref + (Settings.zibcontext.pubLanguage == textLanguage.Multi ? "" : ("(" + Settings.zibcontext.pubLanguage + ")")) + "|" + zref + "]]";
                return "|[[Bestand: " + image + ".png | link="+ zibName.wikiLink(_zibName, Settings.zibcontext.pubLanguage) + "]]||[[" + zibName.wikiLink(_zibName, Settings.zibcontext.pubLanguage) + " |" + zibName.shortName(_zibName) + "]]";
            }

            public override string ToXLS()
            {
                return "[Hyperlink:" + Settings.application.ManualWikiLocation + zibName.wikiLink(Settings.XGenConfig.getLanguageSpecificName(tagValue, Settings.zibcontext.publicatie, Settings.zibcontext.PreReleaseNumber.ToString(), Settings.zibcontext.pubLanguage), Settings.zibcontext.pubLanguage) + "]" + tagNotes;
            }
        }
        // Valuetag interface (tot hier)

        // ===============
        // Concept Class
        // ===============

        public class concept
        {
            public conceptType concepttype;
            public string name;
            public string alias;
            public int id;
            public List<conceptTag> conceptCode;
            public int indentLevel;
            public string definition;
            public List<definitionTag> definitionCodes;
            public string type;
            public string stereotype;
            public string cardinality;
            public List<string> constraint;
            public List<valuesetTag> valuelists;
            public List<authorityTag> authorities;
            public List<zibrefTag> zibrefs;
            public int treepos;
            public int packageId;
            public int maxIndent;

            readonly Dictionary<conceptType, conceptTypeData> cT = new Dictionary<conceptType, conceptTypeData>()
            {
                { conceptType.rootconcept, new conceptTypeData { rowColor = "#E3E3E3", image = "block.png", size = 20 }},        //Zib.png
                { conceptType.container, new conceptTypeData { rowColor = "#E8D7BE", image = "folder.png", size = 16 }},     //Container.png was 20
                { conceptType.reference, new conceptTypeData { rowColor = "transparent", image = "Verwijzing.png", size = 20 }},
                { conceptType.data, new conceptTypeData { rowColor = "transparent", image = "", size = 16 }},
                { conceptType.error, new conceptTypeData { rowColor = "transparent", image = "Multi.png", size = 16 }}
            };


            public override string ToString()
            {
                //            string line = name + "|" + alias + "|" + id + "|" + taglistToString(conceptCode) + "|" + indentLevel + "|" + taglistToString(definitionCodes) + "|" + type + "|" + stereotype + "|" + cardinality + "|" + taglistToString(valuelists) + "|" + taglistToString(zibrefs) + "|" + treepos + "|" + indentLevel + "|" + constraintlistToString(constraint);
                string line = "Positie:" + indentLevel + " | " + packageId + "-" + treepos + ": " + name.PadLeft(name.Length + indentLevel, '-') + "(" + id + "): " + stereotype;
    
                return line;
            }

            public string ToWiki()
            {
                StringBuilder listText = new StringBuilder();
                string _alias;
                string _name;

                // template names bevatten haken waar wiki niet tegen kan
                _alias = Settings.zibcontext.zibPrefix.Contains("template") ? alias.TemplateNameToWiki() : alias;
                _name = Settings.zibcontext.zibPrefix.Contains("template") ? name.TemplateNameToWiki() : name;

                listText.AppendLine("|-style=\"vertical-align:top; background-color: " + cT[concepttype].rowColor + "; \"");
                if (cT[concepttype].image == "") cT[concepttype].image = type + ".png";
                listText.AppendLine("|style = \"text-align:center\" |[[Bestand: " + cT[concepttype].image + "| " + cT[concepttype].size.ToString() + "px | link=]]");
                listText.AppendLine("|" +conceptCode[0].ToWiki());
                //listText.AppendLine("|" + name.PadLeft(name.Length + indentLevel, '>'));
                if (indentLevel>1)
                    for (int i=0 ; i < indentLevel-1; i++)
                            listText.AppendLine("|style = \"width: 7px; padding-left: 0px; padding-right: 0px; border-left: none; border-right: none;\" |");
                if (indentLevel > 0)
                    listText.AppendLine("|style = \"width: 7px; padding-left: 0px; padding-right: 0px; border-left: none; border-right: 1px dotted silver;\" |");
                string arrowImg = concepttype == conceptType.rootconcept || concepttype == conceptType.container ? "arrowdown" : "arrowright";
                listText.AppendLine("|colspan =\"" + (maxIndent + 1- indentLevel).ToString() + "\" style =\"padding-left: 0px\"|<span Id=" + id + " Title=\"" + _alias + "\">" + "[[Bestand: " + arrowImg + ".png | 10px | link=]]" + _name + "</span>");
//                listText.AppendLine("|" + alias);
                listText.AppendLine("|" + cardinality);
                listText.AppendLine("|" + definition);
                listText.AppendLine("|");
                listText.Append((definitionCodes.Count > 0 ? taglistToWiki(definitionCodes) : ""));
                listText.AppendLine("|");
                listText.Append((valuelists.Count > 0 ? taglistToWiki(valuelists) : ""));
                listText.Append((authorities.Count > 0 ? taglistToWiki(authorities) : ""));
                listText.Append((zibrefs.Count > 0 ? taglistToWiki(zibrefs) : ""));
                // Constaints vervallen voor Alias
                // listText.AppendLine("|" + constraintlistToWiki(constraint));

                return listText.ToString();
            }

            public string ToXLS()
            {
                StringBuilder listText = new StringBuilder();
                string tmpName;

                tmpName = name.PadLeft(name.Length + indentLevel, ';');
                listText.Append(tmpName.PadRight(maxIndent - indentLevel + tmpName.Length, ';'));
                listText.Append(";" + (string.IsNullOrEmpty(alias) ? " " : alias));
                listText.Append(";" + (string.IsNullOrEmpty(type) ? " " : type));
                listText.Append(";" + (string.IsNullOrEmpty(cardinality) ? " " : cardinality));
                listText.Append(";" + (string.IsNullOrEmpty(stereotype) ? " " : stereotype));
                listText.Append(";" + conceptCode[0].ToXLS());
                listText.Append(";" + definition.encodeForXLS()); //incell returns
                listText.Append(";" + (definitionCodes.Count > 0 ? taglistToXLS(definitionCodes) : " "));
                listText.Append(";");
                listText.Append((valuelists.Count > 0 ? taglistToXLS(valuelists) : ""));
                listText.Append((authorities.Count > 0 ? taglistToXLS(authorities) : ""));
                listText.Append((zibrefs.Count > 0 ? taglistToXLS(zibrefs) : ""));
                listText.Append(" ");
                listText.AppendLine(";" + constraintlistToXLS(constraint));
                if (concepttype == conceptType.rootconcept || concepttype == conceptType.container)
                    listText.AppendLine("<<RowColor>>" + cT[concepttype].rowColor + "; ; ");
                else if (concepttype == conceptType.container)
                    listText.AppendLine("<<Container>>");
                listText.AppendLine("<<Merge>>" + (indentLevel > 0 ? indentLevel + ";": "") + (maxIndent - indentLevel + 1).ToString());
                return listText.ToString();
            }


            private string taglistToString(IEnumerable<ITag> taglist)
            {
                StringBuilder tagText = new StringBuilder();
                if (taglist.ToList().Count != 0)
                {

                    tagText.Append("<");

                    foreach (ITag t in taglist)
                    {
                        tagText.Append(t.ToString());
                    }
                    tagText.AppendLine(">");
                }
                return tagText.ToString();
            }

            private string taglistToWiki(IEnumerable<ITag> taglist)
            {
                StringBuilder tagText = new StringBuilder();
                if (taglist.ToList().Count != 0)
                {

                    tagText.AppendLine("{|");
                    foreach (ITag t in taglist)
                    {
                        tagText.AppendLine("|-");
                        tagText.AppendLine(t.ToWiki());
                    }
                    tagText.AppendLine("|}");
                }
                return tagText.ToString();

            }

            private string taglistToXLS(IEnumerable<ITag> taglist)
            {
                StringBuilder tagText = new StringBuilder();
                if (taglist.ToList().Count != 0)
                {

                    foreach (ITag t in taglist)
                    {
                        tagText.Append(t.ToXLS());
                        tagText.Append(Settings.xlsSplitCell);   //23-11-2017 tbv meerdere hyperlinks
                    }
                    tagText.Length -= Settings.xlsSplitCell.Length; // haal de laatste xlsSplitCell weg
                }
                // 7-2-2017            return tagText.Length> 0 ? tagText.ToString().Substring(0,tagText.ToString().Length-2): "";
                return tagText.Length > 0 ? tagText.ToString() : "";


            }

            private string constraintlistToString(List<string> constraintlist)
            {
                string line = "";

                if (constraintlist.Count > 0)
                {
                    line = "<";
                    line += string.Join(",", constraintlist);
                    line += ">";
                }
                return line;
            }

            private string constraintlistToWiki(List<string> constraintlist)
            {
                string line = "";

                if (constraintlist.Count > 0)
                {
                    foreach (string constraint in constraintlist)
                        line += constraint + "\r\n";
                    line = line.Substring(0, line.Length - 2);
                }
                return line;
            }

            private string constraintlistToXLS(List<string> constraintlist)
            {
                string line = "";

                if (constraintlist.Count > 0)
                {
                    line += string.Join(",", constraintlist);
                }
                return line;
            }

        }

        // ===============
        // ZibName Class
        // ===============

        static public class zibName

        {
            /// <summary>
            /// Geeft op grond van de naam van de zib de naam zonder versienummer 
            /// </summary>
            /// <param name="zibname"></param>
            /// <returns></returns>
            static public string Name(string zibname)
            {
                return zibname.Substring(0, zibname.IndexOf("-v"));
            }

            /// <summary>
            /// Geeft op grond van de naam van de zib de kale bouwsteennaam zonder prefix of versienummer
            /// </summary>
            /// <param name="zibname"></param>
            /// <returns></returns>
            static public string shortName(string zibname)
            {
                return stripName(zibname);
            }

            /// <summary>
            ///  Geeft de prefix van de zib
            /// </summary>
            /// <param name="zibname"></param>
            /// <returns></returns>
            static public string Prefix(string zibname)
            {
                int endName = zibname.IndexOf("-v");
                int endPrefix = zibname.Substring(0, endName==-1? zibname.Length-1: endName).LastIndexOf('.');
                return endPrefix == -1 ? "" : zibname.Substring(0, endPrefix);
            }


            /// <summary>
            /// Geeft op grond van de naam van de zib alleen het versienummer 
            /// </summary>
            /// <param name="zibname"></param>
            /// <returns></returns>
            static public string Version(string zibname)
            {
                int versionIndex = zibname.IndexOf("-v");
                return versionIndex == -1 ? "??" : zibname.Substring(versionIndex + 2);
            }

            /// <summary>
            /// Hulpmethode om de zibnaam te ontdoen van index en versienummer
            /// </summary>
            /// <param name="zibname"></param>
            /// <returns></returns>
            private static string stripName(string zibname)
            {
                int iStart, iEnd;
                string prefix = Prefix(zibname);
                //iStart = zibname.IndexOf(Settings.zibcontext.zibPrefix);
                iStart = prefix.Length == 0 ? -1 : zibname.IndexOf(prefix);
                iEnd = zibname.IndexOf("-v");
                //if (iStart == -1) iStart = 0; else iStart = iStart + Settings.zibcontext.zibPrefix.Length + 1;
                if (iStart == -1) iStart = 0; else iStart = iStart + prefix.Length + 1;
                if (iEnd == -1) return zibname.Substring(iStart); else return zibname.Substring(iStart, iEnd - iStart);
            }

            /// <summary>
            /// Geeft op grond van de naam van de zib de te gebruiken filenaam voor de wikilinks en -pagina's aan in het format 'naam'-v'versie'('publicatie''taal')
            /// </summary>
            /// <param name="zibName"></param>
            /// <returns></returns>
            public static string wikiLink(string zibName, textLanguage zibLanguage, string publication, bool includePublication = true)
            {
                return shortName(zibName) + "-v" + Version(zibName) + "(" + (includePublication? publication : "") +
                        (zibLanguage == textLanguage.Multi ? "" : zibLanguage.ToString()) + ")";
                //      (zibLanguage == textLanguage.Multi ? "" : Settings.zibcontext.pubLanguage.ToString()) + ")";
            }
            public static string wikiLink(string zibName, textLanguage zibLanguage, bool includePublication = true)
            { 
                return wikiLink(zibName, zibLanguage, Settings.zibcontext.publicatie ?? "??", includePublication);
/*                return shortName(zibName) + "-v" + Version(zibName) + "(" + (Settings.zibcontext.publicatie ?? "??") +
                        (zibLanguage == textLanguage.Multi ? "" : zibLanguage.ToString()) + ")"; */
                //      (zibLanguage == textLanguage.Multi ? "" : Settings.zibcontext.pubLanguage.ToString()) + ")";
            }



            /// <summary>
            /// Geeft op grond van de naam van de zib de te gebruiken filenaam aan in het format 'naam'-v'versie'('taal')
            /// </summary>
            /// <param name="zibName"></param>
            /// <returns></returns>
            public static string fileName(string zibName, textLanguage zibLanguage)
            {
                return Name(zibName) + "-v" + Version(zibName) +
                        (zibLanguage == textLanguage.Multi ? "" : ("(" + zibLanguage.ToString()) + ")");
                //      (zibLanguage == textLanguage.Multi ? "" : ("(" + Settings.zibcontext.pubLanguage.ToString()) + ")");
            }


            /// <summary>
            /// Geeft op grond van de naam van de zib de te gebruiken filenaam aan in het format 'naam'-v'versie'('publicatie''taal')
            /// </summary>
            /// <param name="zibName"></param>
            /// <returns></returns>
            /// 
            public static string fileName2(string zibName, textLanguage zibLanguage)
            {
                return Name(zibName) + "-v" + Version(zibName) +
                        (zibLanguage == textLanguage.Multi ? "" : ("(" + (Settings.zibcontext.publicatie ?? "??") + zibLanguage.ToString()) + ")");
                //      (zibLanguage == textLanguage.Multi ? "" : ("(" + Settings.zibcontext.pubLanguage.ToString()) + ")");
            }
        }
    }
}


