using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Net;
using System.Net.Http;
using System.IO;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using System.Windows.Forms;
using Zibs.Configuration;
using Zibs.ExtensionClasses;

namespace Zibs
{
    namespace ZibExtraction
    {
        public class jira
        {
            private Uri baseurl;
            public Uri BaseUrl
            {
                get { return baseurl; }
                set { baseurl = value; }
            }

            public delegate void NewMessageEventHandler(object sender, EventWithStringArgs e);
            public event NewMessageEventHandler NewMessage;

            protected virtual void OnNewMessage(EventWithStringArgs e)
            {
                if (NewMessage != null)
                {
                    NewMessage(this, e);
                }
            }

            /*
               In hoofdprogramma:
               jira Bits = new jira { BaseUrl = new Uri(Settings.bitscontext.baseURL) };
               Bits.NewMessage += Bits_NewMessage;

                private void Bits_NewMessage(object sender, EventWithStringArgs e)
                {
                    Result.Text += e.Text + "\r\n";
                }
            */
            public void bitsIssues(ref Dictionary<string, string> issueDictionary)
            {
                HttpClient client;

                string user = Settings.bitscontext.bitsUser;
                string passw = Settings.bitscontext.bitsPassword;
                jiraResponse ZibIssues = new jiraResponse();
                string status;

                OnNewMessage(new EventWithStringArgs("Ophalen Bits issues"));
                using (client = createHttpClient())
                {
                    status = getIssues(user, passw, client, ref ZibIssues);
                }
                if (status == "OK")
                {
                    OnNewMessage(new EventWithStringArgs("Inloggen in Bits geslaagd"));

                    getIssueDictionary(ZibIssues, ref issueDictionary);
                    OnNewMessage(new EventWithStringArgs("Issues ingelezen"));

                    // Maak ook de wiki pagina aan met issues
                    issues2Wiki(ZibIssues, Settings.userPreferences.WikiLocation);
                    OnNewMessage(new EventWithStringArgs("Issues wikipagina aangemaakt"));

                    List<string> zibList =  GetZibsWithIssues(ZibIssues);


                }
                else
                    OnNewMessage(new EventWithStringArgs("Issues lezen is foutgegaan. Status: " + status));

            }

            public HttpClient createHttpClient()
            {
                CookieContainer cookieContainer = new CookieContainer();
                HttpClientHandler clientHandler = new HttpClientHandler();
                clientHandler.CookieContainer = cookieContainer;
                clientHandler.UseCookies = true;

                HttpClient client = new HttpClient(clientHandler);
                client.BaseAddress = baseurl;
                client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "Nictiz bot client");
                return client;
            }


            public string getIssues(string user, string passw, HttpClient client, ref jiraResponse ZibIssues)
            {

                string errorResponse = "";
                //                string fields = "key,summary,description,status,customfield_10302,customfield_10301,created,resolutiondate,fixVersions,components";
                string fields = "key,summary,description,status,customfield_10302,customfield_10301,created,resolutiondate,customfield_10801";
                StringBuilder statusField = new StringBuilder();

                // bits query strings voor verschillende output. StatusPubWait wordt nu gebruikt.
                string[] status = (Settings.bitscontext.bitsStatus ?? "").Split(';');
                if (status.Count() == 1 && string.IsNullOrWhiteSpace(status[0]))
                {
                    OnNewMessage(new EventWithStringArgs("Geen issuestatus opgegeven. Ophalen issues gestopt"));
                    return "";
                }
                else
                    foreach (string s in status)
                        if (!string.IsNullOrWhiteSpace(s)) statusField.Append((@"status = """ + s.Trim() + @""" OR ").Replace(" ", "%20"));
                statusField.Length = statusField.Length - 8;
                string statusPubWait = "jql=project%20=%20\"Zorginformatiebouwstenen\"%20AND%20issuetype%20=%20Wijzigingsverzoek%20AND%20(" + statusField.ToString() + ")%20ORDER%20BY%20key%20ASC";

                string payload = baseurl + "search?" + statusPubWait + "&maxResults=1500&fields=" + fields;

                string Content = "";
                HttpContent httpContent;
                httpContent = new StringContent(Content);
                HttpResponseMessage response = null;
                string responseStatus = "";

                client.Timeout = new TimeSpan(0, 0, 30);

                var byteArray = Encoding.ASCII.GetBytes(user + ":" + passw.Decode());
                var header = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
                client.DefaultRequestHeaders.Authorization = header;
                try
                {
                    response = client.GetAsync(payload).Result;
                    if (!response.IsSuccessStatusCode)
                    {
                        errorResponse = response.Content.ReadAsStringAsync().Result; // Fails with ObjectDisposedException
                    }
                    _= response.EnsureSuccessStatusCode();
                }
                catch (Exception ex) 
                {
                    string error = ex.GetBaseException().Message ?? "";
                    if (!string.IsNullOrWhiteSpace(error) && ex.InnerException != null) error += "\r\n" + ex.InnerException.Message;
                    if (!string.IsNullOrWhiteSpace(errorResponse)) error += (string.IsNullOrWhiteSpace(error) ? "" : "\r\n") + "Jira meldt: "+ errorResponse;
                    MessageBox.Show(error); 
                }
                if (response == null)
                {
                    responseStatus = "Failure";
                }
                else 
                { 
                    if(response.StatusCode == HttpStatusCode.OK)
                    {
                        responseStatus = response.StatusCode.ToString();
                        string responseString = response.Content.ReadAsStringAsync().Result;
                        ZibIssues = JsonConvert.DeserializeObject<jiraResponse>(responseString);
                        OnNewMessage(new EventWithStringArgs("Zib issues ingelezen, aantal: " + ZibIssues.total));
                    }
                    else
                    {
                        OnNewMessage(new EventWithStringArgs("Fout bij inlezen issues uit Bits: " + response.StatusCode.ToString() + " - " + response.ReasonPhrase.ToString()));
                    }
                }
                return responseStatus;
            }

            public List<string> GetZibsWithIssues(jiraResponse zibs)
            {
                List<string> zibsWithIssues = null;
                List<string> isuesWithoutZib = new List<string>();
                Dictionary<string, string[]> issuesForReleaseNotes = new Dictionary<string, string[]>();
                Dictionary<string, List<string>> zibsAndIssues = new Dictionary<string, List<string>>();

                foreach (jiraResponse.jiraIssue bitsIssue in zibs.issues)
                {
                    string issueNo = bitsIssue.key;
                    var zibNames = bitsIssue.fields.customfield_10301;
                    if (zibNames != null)
                    {
                        foreach (jiraResponse.jiraIssue.jiraFields.jiraCF10301 zibName in zibNames)
                        {
                            if (!zibsAndIssues.ContainsKey(zibName.value))
                            {
                                zibsAndIssues.Add(zibName.value, new List<string> { issueNo });
                            }
                            else
                            {
                                if (!zibsAndIssues[zibName.value].Contains(issueNo))
                                    zibsAndIssues[zibName.value].Add(issueNo);
                            }
                        }
                    }
                    else
                        isuesWithoutZib.Add(issueNo);

                    if (!issuesForReleaseNotes.ContainsKey(issueNo))
                    {
                        var zibList = bitsIssue.fields.customfield_10301 == null ? "" : string.Join(",", (bitsIssue.fields.customfield_10301.Select(x => x.value)).ToArray());
                        string[] issueInfo = { zibList, bitsIssue.fields.customfield_10302 ?? "" };
                        issuesForReleaseNotes.Add(issueNo, issueInfo);
                    }
                }

                zibsWithIssues = zibsAndIssues.Select(x => x.Key).OrderBy(x=>x).ToList();

               return zibsWithIssues;
            }



            public void issues2Wiki(jiraResponse zibs, string path)
            {

                string wikiPath = Path.Combine(path, @"ZIBIssues.wiki");
                StringBuilder sw = new StringBuilder();
                DateTime createDate, dateResolved;

                string issueTitle;
                bool isLink;
                string bitsUrl = "";
                int apiStart = Settings.bitscontext.bitsBaserurl.IndexOf("rest");
                if (apiStart != -1)
                {
                    bitsUrl = Settings.bitscontext.bitsBaserurl.Substring(0, apiStart - 1);
                    isLink = true;
                }
                else
                    isLink = false;

                    sw.AppendLine("__NOTOC__");
                for (int i = 0; i < zibs.total; i++)
                {
                    //4-12-2021 Titel omgezet in link naar Bits
                    if (isLink)
                    {
                        issueTitle = "[" + bitsUrl + "/browse/" + zibs.issues[i].key + " " + zibs.issues[i].key + "]";
                    }
                    else
                    {
                        issueTitle = zibs.issues[i].key;
                    }

                    sw.AppendLine("==" + issueTitle + "=="); //==\n
                    sw.AppendLine("'''" + zibs.issues[i].fields.summary + "'''"); // <br>\n
                    sw.AppendLine("{|style=\"font-size:90%; width:750px\""); // \n
                    sw.AppendLine("|-style=\"vertical-align:top\"\n|style=\"width:180px\"|Aangemaakt op: \n|style=\"width:300px\"|" + (DateTime.TryParse(zibs.issues[i].fields.created, out createDate) ? createDate.ToString("dd-MM-yyyy") : zibs.issues[i].fields.created));  // + "\n"
                    sw.AppendLine("|style=\"width:120px\"|Status: \n|" + zibs.issues[i].fields.status.name); // \n
                    /*
                                        sw.Append("|Component:\n|"); //AppendLine
                                        int iComp = zibs.issues[i].fields.components.Count();
                                        if (iComp > 0)
                                        {
                                            for (int j = 0; j < iComp; j++)
                                            {
                                                sw.AppendLine(zibs.issues[i].fields.components[j].name + "<br>"); //"<br>\n"
                                            }
                                            sw.Length = sw.Length - 6; // verwijder de laatste <br>
                                            sw.AppendLine(); // maar voeg de newline weer toe
                                        }
                    */
                    sw.Append("|-style=\"vertical-align:top\"\n"); // \n
                    sw.Append("|Onderdeel van: \n|"); // \n
                    int? iReleases = zibs.issues?[i].fields.customfield_10801?.Count();
                    if (iReleases > 0)
                    {
                        sw.AppendLine(zibs.issues[i].fields.customfield_10801[0].value); 
                    }

                    string strDateResolved = DateTime.TryParse(zibs.issues[i].fields.resolutiondate, out dateResolved) ? dateResolved.ToString("dd-MM-yyyy") : zibs.issues[i].fields.resolutiondate;
                    sw.AppendLine("|Publicatiedatum: \n|" + strDateResolved ?? "" + "\n");

                    sw.Append("|-style=\"vertical-align:top\"\n|Het betreft de bouwstenen:\n|"); //AppendLine
                    int? iZibs = zibs.issues?[i].fields.customfield_10301?.Count();
                    if (iZibs > 0)
                    {
                        for (int k = 0; k < iZibs; k++)
                        {
                            sw.AppendLine("'''" + zibs.issues[i].fields.customfield_10301[k].value + "'''<br>"); //was value 16-09-2017 + <br>\n
                        }
                        sw.Length = sw.Length - 6; // verwijder de laatste <br>
                        sw.AppendLine(); // maar voeg de newline weer toe
                    }
                    else
                        sw.AppendLine();
                    /*
                                        sw.Append("|Publicatieversie: \n|");
                                        if (zibs.issues[i].fields.fixVersions.Count() > 0)
                                        {
                                            int iMax = zibs.issues[i].fields.fixVersions.ToList().IndexOf(zibs.issues[i].fields.fixVersions.Max());
                                            sw.Append(zibs.issues[i].fields.fixVersions[iMax].name);
                                        }
                    */
                    sw.AppendLine("|}"); // \n
                    sw.AppendLine("------"); //\n
                    sw.AppendLine("'''Omschrijving: ''' <br>\n<nowiki>" + zibs.issues[i].fields.description + "</nowiki><br>"); // \n
                    sw.AppendLine("'''Besluit: ''' <br>\n<nowiki>" + (zibs.issues[i].fields.customfield_10302 ?? "") + "</nowiki><br><br>");  // \n

                }
                sw.AppendLine("<BR>\n------\n");
                sw.AppendLine("Deze wiki pagina is gegenereerd op " + DateTime.Now.ToString(new CultureInfo("nl-NL")));

                File.WriteAllText(wikiPath, sw.ToString());

            }

            public void getIssueDictionary(jiraResponse zibs, ref Dictionary<string, string> issueDictionary)
            {
                for (int i = 0; i < zibs.issues?.Count(); i++)
                {
                    issueDictionary.Add(zibs.issues[i].key, zibs.issues[i].fields.summary);
                }
            }
        }



        public class jiraResponse
        {
            public int total;
            public jiraIssue[] issues;


            public class jiraIssue
            {
                public int id;
                public string key;
                public jiraFields fields;

                public class jiraFields
                {
                    public string summary;
                    public jiraStatus status;
                    public string created; 
                    public string customfield_10302;                // ReleaseNotes
                    public string description;
                    public jiraCF10301[] customfield_10301;         // bouwstenen
//                    public jiraFixVersion[] fixVersions;
//                    public jiraComponent[] components;
                    public jiraCF10801[] customfield_10801;         // (pre)publicatie
                    public string resolutiondate;

                    public class jiraStatus
                    {
                        public string name;
                    }

                    public class jiraCF10301
                    {
                        public string value; //was name 16-09-2017
                    }

                    /*
                    public class jiraFixVersion
                    {
                        public string name;
                    }

                    public class jiraComponent
                    {
                        public string name;
                    }
                    */

                    public class jiraCF10801
                    {
                        public string value;
                    }
                }
            }
        }
    }
}