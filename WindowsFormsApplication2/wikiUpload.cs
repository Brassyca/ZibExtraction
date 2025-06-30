using System;
using System.Collections.Generic;
using System.Text;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Net.Http;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Windows.Forms;
using System.Linq;
using Zibs.Configuration;
using Zibs.tekstManager;
using Zibs.ExtensionClasses;


namespace Zibs
{
    namespace ZibExtraction
    {
        public class Wiki
        {
            private Uri baseurl;
            private string requesturl;

            #region Public properties

            public Uri BaseUrl
            {
                get { return baseurl; }
                set { baseurl = value; }
            }

            public string RequestUrl
            {
                get { return requesturl; }
                set { requesturl = value; }
            }
            #endregion

            #region Public event
            public delegate void NewMessageEventHandler(object sender, ErrorMessageEventArgs e);
            public event NewMessageEventHandler NewMessage;

            protected virtual void OnNewMessage(ErrorMessageEventArgs e)
            {
                if (NewMessage != null)
                {
                    NewMessage(this, e);
                }
            }
            #endregion

            /* in hoofdmodule

                string adres = Settings.wikicontext.baseUrl; 
                string apiurl = "api.php";
                Wiki zibWiki = new Wiki { BaseUrl = new Uri(adres), RequestUrl = apiurl };
                zibWiki.NewMessage += zibWiki_NewMessage;

                private void zibWiki_NewMessage(object sender, ErrorMessageEventArgs e)
                {
                    Result.Text += e.Text + "\r\n";
                }


    */
            public void uploadWikiAll(string purgeList)
            {
                string editToken;
                int section;
                HttpClient client;

                string user = Settings.wikicontext.wikiUser;
                string passw = Settings.wikicontext.wikiPassword;

                StringBuilder mainPages = new StringBuilder();  // haal de namen van de hoofdpagina's op
                textManager tm = new textManager("ZibExtractionLabels.cfg");
                if (tm.dictionaryOK)
                {
                    foreach (textLanguage value in Enum.GetValues(typeof(textLanguage)))
                        if (value != textLanguage.Multi)
                        {
                            tm.Language = value;
                            mainPages.Append(tm.getLabel(Settings.wikicontext.MainPage) + "|");
                            for (int i = 2015; i <= DateTime.Now.Year; i++)
                                mainPages.Append(tm.getLabel("wikiReleasePage") + "_" + i.ToString() + "(" + value.ToString() + ") | ");
                        }
                    mainPages.Length -= 3;
                }

                using (client = createHttpClient())
                {
                    bool loggedIn = login(user, passw, client, out editToken);
                    if (loggedIn)
                    {
                        if (Directory.Exists(Settings.userPreferences.DiagramLocation))
                        {
                            OnNewMessage(new ErrorMessageEventArgs(ErrorType.general, "Upload diagrammen"));
                            string[] imageEntries = Directory.GetFiles(Settings.userPreferences.DiagramLocation);
                            foreach (string imageName in imageEntries)
                            {
                                uploadImage(imageName, client, editToken);
                                string name = Path.GetFileName(imageName);
                                int i = name.IndexOf(".");
                                purgePage("Bestand:" + Path.GetFileName(name), client);
                            }
                        }
                        //                  upload ook pdf en xlsx bestanden van de ZIB voor download
                        if (Directory.Exists(Settings.userPreferences.ZIB_RTFLocation))
                        {
                            OnNewMessage(new ErrorMessageEventArgs(ErrorType.general, "Upload pdf's"));

                            string[] pdfEntries = Directory.GetFiles(Settings.userPreferences.ZIB_RTFLocation, "*.pdf");
                            foreach (string pdfName in pdfEntries)
                            {
                                uploadImage(pdfName, client, editToken);
                                purgePage("Bestand:" + Path.GetFileName(pdfName), client);
                            }
                        }

                        if (Directory.Exists(Settings.userPreferences.XLSLocation))
                        {
                            OnNewMessage(new ErrorMessageEventArgs(ErrorType.general, "Upload spreadsheets"));
                            string[] xlsEntries = Directory.GetFiles(Settings.userPreferences.XLSLocation);
                            foreach (string xlsName in xlsEntries)
                            {
                                uploadImage(xlsName, client, editToken);
                                purgePage("Bestand:" + Path.GetFileName(xlsName), client);
                            }
                        }

                        if (Directory.Exists(Settings.userPreferences.WikiLocation))
                        {
                            OnNewMessage(new ErrorMessageEventArgs(ErrorType.general, "Upload wiki pagina's"));
                            string[] fileEntries = Directory.GetFiles(Settings.userPreferences.WikiLocation);
                            foreach (string fileName in fileEntries)
                            {
                                // Kijk of de filenaam de string "_Section" bevat ten teken dat er slechts een sectie vervangen wordt
                                string pageName = getPageName(Path.GetFileNameWithoutExtension(fileName), out section);
                                string extension = Path.GetExtension(fileName);
                                // voeg evt. indien gewenst namespaces als extensie toe om page upload hiervan mogelijk te maken
                                // doe dit niet voor beelden en bestanden: deze worden 'uploadImage'verstuurd: dit kan grotere bestanden aan
                                switch (extension)
                                {
                                    case ".wiki":
                                        break;
                                    case ".template":
                                        pageName = "Sjabloon:" + pageName;
                                        break;
                                    case ".special":
                                        pageName = "Speciaal:" + pageName;
                                        break;
                                    case ".category":
                                        pageName = "Categorie:" + pageName;
                                        break;
                                    case ".mediawiki":
                                        pageName = "Mediawiki:" + pageName;
                                        break;
                                    default:
                                        pageName = "";
                                        break;
                                }
                                if (pageName != "")
                                {
                                    if (section == 0)
                                        uploadPage(pageName, fileName, client, editToken);
                                    else
                                        uploadPage(pageName, section, fileName, client, editToken);
                                    purgePage(pageName, client);
                                }
                            }
                        }
                        OnNewMessage(new ErrorMessageEventArgs(ErrorType.general, "Purging wiki pagina's"));
                        if (purgeList.Length > 0) purgePage(purgeList, client);
                        purgePage(mainPages.ToString(), client);
                    }
                }

            }

            private string getPageName(string filename, out int section)
            {
                section = 0;
                int j = filename.IndexOf("_Section");
                if (j != -1)
                    if (int.TryParse(filename.Substring(j + 8, 1), out section)) // gaat fout bij meer dan 9 secties
                        filename = filename.Substring(0, j);
                filename = filename.Replace('$', ':');
                return filename;
            }



            public bool Validator(object sender, X509Certificate certificate, X509Chain chain,
                                      SslPolicyErrors sslPolicyErrors)
            {
                // If the certificate is a valid, signed certificate, return true.
                if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
                {
                    return true;
                }

                // If there are errors in the certificate chain, look at each error to determine the cause.
                if ((sslPolicyErrors & System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) != 0)
                {
                    if (chain != null && chain.ChainStatus != null)
                    {
                        foreach (System.Security.Cryptography.X509Certificates.X509ChainStatus status in chain.ChainStatus)
                        {
                            if ((certificate.Subject == certificate.Issuer) &&
                               (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot))
                            {
                                // Self-signed certificates with an untrusted root are valid. 
                                continue;
                            }
                            else
                            {
                                if (status.Status != System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError)
                                {
                                    // If there are any other errors in the certificate chain, the certificate is invalid,
                                    // so the method returns false.
                                    if (Settings.ignoreCertificateErrors == false)
                                    {
                                        DialogResult answ = MessageBox.Show("De volgende certificaat fouten zijn geconstateerd: " + sslPolicyErrors.ToString() + "\r\n" +
                                            "Certificaat onderwerp: " + certificate.Subject + "\r\n" +
                                            "Uitgever: " + certificate.Issuer + "\r\n" +
                                            "Expiratie datum: " + certificate.GetExpirationDateString() + "\r\n" +
                                            "Wil je toch doorgaan?",
                                            "Beveiligingsfout", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                                        if (answ == DialogResult.Yes)
                                            return true;
                                        else
                                            return false;
                                    }
                                    else
                                        return true;  // gebruiker heeft in configuratie gekozen certificaar fouten te negeren
                                                      // hoort sowieso false te zijn, maar tijdelijk mogelijk true omdat het cerfificaat van testzibs niet klopt.
                                }
                            }
                        }
                    }

                    // When processing reaches this line, the only errors in the certificate chain are 
                    // untrusted root errors for self-signed certificates. These certificates are valid
                    // for default Exchange server installations, so return true.
                    return true;
                }
                else
                {
                    // In all other cases, return false.
                    return false;
                }
            }


            public HttpClient createHttpClient()
            {
                CookieContainer cookieContainer = new CookieContainer();
                HttpClientHandler clientHandler = new HttpClientHandler();
                clientHandler.CookieContainer = cookieContainer;
                clientHandler.UseCookies = true;

                ServicePointManager.ServerCertificateValidationCallback = Validator;


                HttpClient client = new HttpClient(clientHandler);
                client.BaseAddress = baseurl;
                client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "Nictiz-ZIB2Wiki-client");
                return client;
            }

            public bool login(string user, string passw, HttpClient client, out string editToken)
            {
                HttpContent httpContent;
                HttpResponseMessage response = null;
                string Content;
                string responseString = "";
                string token = "";
                bool loggedIn = false;

                string login = "?action=login&format=json";
                string getlogintoken = "?action=query&meta=tokens&type=login&format=json";
                string getedittoken = "?action=query&meta=tokens&format=json&continue=";

                // JSON response parse definities
                var loginResponse = new { query = new { tokens = new { logintoken = string.Empty } } };
                var login2Response = new { login = new { result = string.Empty, lguserid = 0, lgusername = string.Empty, lgtoken = string.Empty, cookieprefix = string.Empty, sessionid = string.Empty } };
                var editTokens = new { batchcomplete = string.Empty, query = new { tokens = new { csrftoken = string.Empty } } };

                Content = "";
                httpContent = new StringContent(Content);

                // login; nieuw haal eerst een logintoken op

                try
                {
                    response = client.GetAsync(requesturl + getlogintoken).Result;
                    response.EnsureSuccessStatusCode();

                    responseString = response.Content.ReadAsStringAsync().Result;
                    if (responseString.Contains("error"))
                    {
                        reportJsonError(responseString);
                    }
                    else
                    {
                        var JsonResponse1 = JsonConvert.DeserializeAnonymousType(responseString, loginResponse);
                        token = JsonResponse1.query.tokens.logintoken;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.GetBaseException().Message);
                    OnNewMessage(new ErrorMessageEventArgs(ErrorType.error, "Upload afgebroken wegens fouten"));
                    editToken = "";
                    return false;
                }

                // login fase 2

                //                Content = "";
                //                httpContent = new StringContent(Content);
                StringBuilder sb = new StringBuilder();
                sb.Append("lgname=" + user);
                sb.Append("&lgpassword=" + WebUtility.UrlEncode(passw.Decode()));
                sb.Append("&lgtoken=" + WebUtility.UrlEncode(token));
                httpContent = new ByteArrayContent(Encoding.UTF8.GetBytes(sb.ToString()));
                httpContent.Headers.TryAddWithoutValidation("content-type", "application/x-www-form-urlencoded");
                try
                {
                    response = client.PostAsync(requesturl + login, httpContent).Result;
                    response.EnsureSuccessStatusCode();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.GetBaseException().Message);
                    OnNewMessage(new ErrorMessageEventArgs(ErrorType.error, "Upload afgebroken wegens fouten"));
                    editToken = "";
                    return false;
                }

                responseString = response.Content.ReadAsStringAsync().Result;
                if (responseString.Contains("error"))
                {
                    reportJsonError(responseString);
                }
                else
                {

                    var JsonResponse1 = JsonConvert.DeserializeAnonymousType(responseString, login2Response);

                    //Console.WriteLine("Login2: " + response.StatusCode);
                    OnNewMessage(new ErrorMessageEventArgs(ErrorType.information, "Wiki Login: " + JsonResponse1.login.result));
                }

                // get editToken

                try
                {
                    response = client.GetAsync(requesturl + getedittoken).Result;
                    response.EnsureSuccessStatusCode();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.GetBaseException().Message);
                    OnNewMessage(new ErrorMessageEventArgs(ErrorType.error, "Upload afgebroken wegens fouten"));
                    editToken = "";
                    return false;
                }

                responseString = response.Content.ReadAsStringAsync().Result;
                if (responseString.Contains("error"))
                {
                    reportJsonError(responseString);
                    editToken = "";
                }
                else
                {
                    var JsonResponse2 = JsonConvert.DeserializeAnonymousType(responseString, editTokens);
                    editToken = JsonResponse2.query.tokens.csrftoken;
                    //                OnNewMessage(new ErrorMessageEventArgs("Edittoken: " + editToken));
                    loggedIn = true;
                }

                return loggedIn;
            }

            // Multipart image upload

            public void uploadImage(string fileName, HttpClient client, string editToken)
            {
                Dictionary<string, string> mimeTypes = new Dictionary<string, string>
                {
                    {".bmp", "image/bmp"},
                    {".gif", "image/gif"},
                    {".ico", "image/x-icon"},
                    {".jpe", "image/jpeg"},
                    {".jpeg", "image/jpeg"},
                    {".jpg", "image/jpeg"},
                    {".png", "image/png"},
                    {".pdf", "application/pdf"},
                    {".tif", "image/tiff"},
                    {".tiff", "image/tiff"},
                    //{".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}, //mime type aangepast tbv MediaWiki
                    {".xml", "application/xml"},
                    {".xmi", "application/xml"},
                    {".docx", "application/zip"},
                    {".xlsx", "application/zip"},
                    {".zip", "application/zip"}
                };
                string errorResponse = "";
                StringBuilder sb = new StringBuilder();
                FileInfo fi = new FileInfo(fileName);
                ErrorType eType = ErrorType.information;
                if (mimeTypes.ContainsKey(fi.Extension))
                {

                    HttpResponseMessage response = null;
                    string responseString;

                    // JSON response parse definities
                    //               var imageUploadResult = new { upload = new { result = string.Empty, filename = string.Empty, warnings = new { badfilename = string.Empty } } };

                    //  Multipart

                    MultipartFormDataContent form = new MultipartFormDataContent("----Zib Upload---" + DateTime.Now.Ticks.ToString("x"));
                    byte[] byteArray = File.ReadAllBytes(fi.FullName);
                    ByteArrayContent bc = new ByteArrayContent(byteArray);
                    bc.Headers.Add("Content-Type", mimeTypes[fi.Extension]);
                    form.Add(new StringContent("upload"), "\"action\"");
                    form.Add(new StringContent("bot"), "\"assert\"");
                    form.Add(new StringContent("json"), "\"format\"");
                    form.Add(bc, "\"file\"", "\"" + fi.Name + "\"");
                    //                form.Add(new StringContent(fileName), "filename");
                    form.Add(new StringContent(fi.Name), "\"filename\"");
                    form.Add(new StringContent("bot_upload dd." + DateTime.Now.ToString()), "\"comment\"");
                    form.Add(new StringContent("yes"), "\"ignorewarnings\"");
                    form.Add(new StringContent(editToken), "\"token\"");

                    try
                    {
                        response = client.PostAsync(requesturl, form).Result;
                        response.EnsureSuccessStatusCode();
                        if (!response.IsSuccessStatusCode)
                        {
                            errorResponse = response.Content.ReadAsStringAsync().Result; // Fails with ObjectDisposedException
                        }
                    }
                    catch (Exception ex)
                    {
                        errorResponse += (string.IsNullOrEmpty(errorResponse) ? "" : "\r\n") + ex.GetBaseException().Message;
                        sb.AppendLine(errorResponse);
                        sb.AppendLine("Upload van " + fileName + " wegens fouten afgebroken");
                        OnNewMessage(new ErrorMessageEventArgs(ErrorType.error, sb.ToString()));
                        return;
                    }

                    responseString = response.Content.ReadAsStringAsync().Result;
                    JObject JResult = JObject.Parse(responseString);
                    string result = JResult["upload"]?.Value<string>("result");
                    if (result == "Success")
                    {
                        var file = JResult["upload"]?.Value<string>("filename");
                        if (file == null) eType = ErrorType.error;
                        sb.Append("Upload image: " + (file != null ? file : "File fout!!") + " : " + result);

                        var parts = JResult["upload"]["warnings"]?.Children();
                        if (parts != null)
                        {
                            sb.Append(" met waarschuwing: ");
                            foreach (JToken part in parts)
                            {
                                sb.Append(((JProperty)part).Name + " ");
                            }
                        }
                    }
                    else
                    {
                        eType = ErrorType.error;
                        var errorMessage = JResult["error"]?.Value<string>("code");
                        if (errorMessage == "fileexists-no-change")
                        {
                            string shortName = fileName.Substring(fileName.LastIndexOf("\\") == -1 ? 0 : (fileName.LastIndexOf("\\") + 1));
                            sb.Append("No upload of image: " + shortName + " : " + errorMessage);
                            eType = ErrorType.information;
                        }
                        else
                            sb.Append("Onverwacht resultaat: " + fileName + " :\r\nresult: " + JResult.ToString());
                    }
                    response.Dispose();
                    form.Dispose();
                }
                else
                {
                    eType = ErrorType.error;
                    sb.Append("Upload bestand " + fileName + "niet uitgevoerd: niet ondersteund filetype: " + fi.Extension);
                }
                OnNewMessage(new ErrorMessageEventArgs(eType, sb.ToString()));
            }
            public void uploadPage(string title, int section, string fileName, HttpClient client, string editToken) 
            {
                string pageDate = File.ReadAllText(fileName);
                uploadPageData(title, section, pageDate, client, editToken);
            }


            public void uploadPageData(string title, int section, string text, HttpClient client, string editToken)
            {
                HttpContent httpContent; //= new StringContent(""); ;
                HttpResponseMessage response = null;
                string responseString;
                string errorResponse = "";

                // JSON response parse definities
                var pageUploadResult = new { edit = new { result = string.Empty, pageid = 0, title = string.Empty, contentmodel = string.Empty, nochange = string.Empty } };

                StringBuilder sb = new StringBuilder();
                sb.Append("action=edit");
                sb.Append("&assert=bot");
                sb.Append("&format=json");
                sb.Append("&title=" + WebUtility.UrlEncode(title));
                if (section != 0) sb.Append("&section=" + section.ToString());
                //string text = File.ReadAllText(fileName);
                sb.Append("&text=" + WebUtility.UrlEncode(text));
                sb.Append("&token=" + WebUtility.UrlEncode(editToken));

                httpContent = new ByteArrayContent(Encoding.UTF8.GetBytes(sb.ToString()));
                httpContent.Headers.TryAddWithoutValidation("content-type", "application/x-www-form-urlencoded");

                try
                {
                    response = client.PostAsync(requesturl, httpContent).Result;
                    response.EnsureSuccessStatusCode();
                    if (!response.IsSuccessStatusCode)
                    {
                        errorResponse = response.Content.ReadAsStringAsync().Result; // Fails with ObjectDisposedException
                    }
                }
                catch (Exception ex)
                {
                    errorResponse += string.IsNullOrEmpty(errorResponse) ? "" : "\r\n" + ex.GetBaseException().Message;
                    errorResponse += "\r\nUpload van " + title + " wegens fouten afgebroken";
                    OnNewMessage(new ErrorMessageEventArgs(ErrorType.error, errorResponse));
                    return;
                }
                responseString = response.Content.ReadAsStringAsync().Result;
                JObject JResult = JObject.Parse(responseString);
                string result = JResult["edit"]?.Value<string>("result");
                if (result == "Success")
                {
                    var pageTitle = JResult["edit"]?.Value<string>("title");
                        
//                       var JsonResponse = JsonConvert.DeserializeAnonymousType(responseString, pageUploadResult);
                    OnNewMessage(new ErrorMessageEventArgs(pageTitle == null? ErrorType.error : ErrorType.information, "Upload page: " + (pageTitle != null ? title : "Pagina fout!!") + ": " + result));
                //Console.WriteLine(responseString);
                }
                else  
                {
                    OnNewMessage(new ErrorMessageEventArgs(ErrorType.error, "Onverwacht resultaat: " + title + " :\r\nresult: " + JResult.ToString()));
                }
                response.Dispose();
                httpContent.Dispose();
            }

            public void purgePage(string title, HttpClient client)
            {
                HttpContent httpContent; //= new StringContent(""); ;
                HttpResponseMessage response = null;
                string responseString;
                JProperty resultPart;

                int i;

                // JSON response parse definities
//                var pagePurgeResult = new { batchcomplete = string.Empty, purge = new[] { new { ns = 0, title = string.Empty, string.Empty } } };
//                var pagePurgeResult = new { batchcomplete = string.Empty, purge = new List<string>()};

                StringBuilder sb = new StringBuilder();
                sb.Append("action=purge");
                sb.Append("&forcelinkupdate");
                sb.Append("&assert=bot");
                sb.Append("&format=json");
                sb.Append("&titles=" + WebUtility.UrlEncode(title));
                httpContent = new ByteArrayContent(Encoding.UTF8.GetBytes(sb.ToString()));
                httpContent.Headers.TryAddWithoutValidation("content-type", "application/x-www-form-urlencoded");
                try
                {
                    response = client.PostAsync(requesturl, httpContent).Result;
                    response.EnsureSuccessStatusCode();


                    responseString = response.Content.ReadAsStringAsync().Result;
                    if (responseString.Contains("error"))
                    {
                        reportJsonError(responseString);
                    }
                    else
                    {
                        JObject JResult = JObject.Parse(responseString);
                        JArray purgeResults = JResult["purge"].Value<JArray>();
                        foreach (JObject purgeResult in purgeResults)
                        {
                            sb.Clear();
                            sb.Append(JResult.Properties().First().Name + ": " + JResult.Properties().First().Value + "; ");
                            sb.Append(purgeResult["title"].ToString() + ": ");
                            var p = purgeResult.Children<JProperty>();
                            for (i = 2; i < p.Count(); i++)
                            {
                                resultPart = p.ElementAt<JProperty>(i);
                                sb.Append(resultPart.Name + " ");
                            }
                            OnNewMessage(new ErrorMessageEventArgs(ErrorType.information, "Purge page: " + sb.ToString()));
                        }

                        //                        var JsonResponse = JsonConvert.DeserializeAnonymousType(responseString, pagePurgeResult);
                        //                        for (i = 0; i < JsonResponse.purge.Count; i++)
                        //                            OnNewMessage(new ErrorMessageEventArgs("Purge page " + ": " + JsonResponse.purge[i].ToString()));
                    }
                }
                catch (Exception ex)
                {
//                    MessageBox.Show(ex.GetBaseException().Message);
                    OnNewMessage(new ErrorMessageEventArgs(ErrorType.error, "Purge page " + title + " :" + ex.GetBaseException().Message));
                }

                response.Dispose();
                httpContent.Dispose();
            }


            public void uploadPage(string title, string fileName, HttpClient client, string editToken)
            {
                int sectie = 0;
                uploadPage(title, sectie, fileName, client, editToken);
            }

            public void reportJsonError(string responseString)
            {
                var responseError = new { error = new { code = string.Empty, info = string.Empty } };
                var JsonResponse = JsonConvert.DeserializeAnonymousType(responseString, responseError);
                OnNewMessage(new ErrorMessageEventArgs(ErrorType.error, "Upload page: " + JsonResponse.error.code + ": " + JsonResponse.error.info));
            }
        }
    }
}