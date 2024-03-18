using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Web;
using System.Threading.Tasks;
using System.Windows;

namespace DragAndDrop4Outlook
{
    class KLICApi
    {
        public async Task<string> getObjectID(string token, string url, string KlicNummer)
        {
            try
            {
                HttpClient http = new HttpClient();
                var nvc = new Dictionary<string, string>();
                nvc.Add("f", "pjson");
                nvc.Add("token", token);
                nvc.Add("returnIdsOnly", "true");
                nvc.Add("where", "Klic_Meldnummer = \'" + KlicNummer + "\'");
                MultipartFormDataContent content = FormUrlEncodedContentWorkaround(nvc);
                var response = await http.PostAsync(url, content);
                //(url + "?token=" + token, content);
                var result = response.Content.ReadAsStringAsync().Result;
                JObject jsonObject = JObject.Parse(result);
                if (jsonObject["objectIds"] != null)
                {
                    var objectId = jsonObject["objectIds"].First.Value<string>();
                    return objectId;
                }
                else
                {
                    if (!isTokenValid(jsonObject)) return null;
                }


                return "";
            }
            catch
            {

                return "";
            }
        }
        public bool isTokenValid(JObject jsonObject)
        {
            if (jsonObject["error"] != null) if (jsonObject["error"]["code"] != null)
                {
                    if (jsonObject["error"]["code"].ToString().Equals("498"))
                    {
                        string msg = jsonObject["error"]["message"].ToString();
                        MessageBox.Show("De token is niet langer geldig. Probeer opnieuw in te loggen.");
                        return false;
                    }
                }
            return true;
        }
        public async Task<bool> addAttachment(string token, string url, string filePath)
        {
            try
            {
                using (HttpClient http = new HttpClient())
                {
                    var nvc = new Dictionary<string, string>();
                    nvc.Add("f", "json");
                    nvc.Add("token", token);
                    MultipartFormDataContent content = FormUrlEncodedContentWorkaround(nvc);
                    string fileName = System.IO.Path.GetFileName(filePath);
                    FileStream fs = System.IO.File.OpenRead(filePath);
                    using (var br = new BinaryReader(fs))
                    {
                        byte[] data = br.ReadBytes((int)fs.Length);
                        br.Dispose();
                        ByteArrayContent bytes = new ByteArrayContent(data, 0, data.Count());
                        data = null;
                        bytes.Headers.ContentType = MediaTypeHeaderValue.Parse(MimeMapping.MimeUtility.GetMimeMapping(fileName));
                        content.Add(bytes, "attachment", fileName);
                        var response = await http.PostAsync(url + "?token=" + token, content);
                        string resp = response.Content.ReadAsStringAsync().Result;
                        fs.Close();
                        return resp != null;
                    }
                    //logger.LogFile(resp);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return false;
            }
        }
        public async Task<string> getToken(string username, string psw, string portalUrl)
        {
            try
            {
                HttpClient http = new HttpClient();
                var nvc = new Dictionary<string, string>();
                nvc.Add("f", "json");
                nvc.Add("username", username);
                nvc.Add("password", psw);
                nvc.Add("client", "referer");
                nvc.Add("referer", portalUrl);
                nvc.Add("expiration", "1440");
                MultipartFormDataContent content = FormUrlEncodedContentWorkaround(nvc);
                var response = await http.PostAsync(portalUrl + "sharing/rest/generateToken", content);
                JObject jsonObject = JObject.Parse(await response.Content.ReadAsStringAsync());
                string token = jsonObject.Value<string>("token");
                return token;
            }
            catch (Exception exception)
            {
                return "";
            }
        }
        public void getKLICMeldingItems()
        {
        }
        public void getKLICMeldingItem()
        {
        }
        public bool FindKLICMeldingItem()
        {
            return false;
        }
        public void addAttachemtToKLICMElding()
        {
        }
        public MultipartFormDataContent FormUrlEncodedContentWorkaround(IEnumerable<KeyValuePair<string, string>> nameValueCollection)
        {
            var content = new MultipartFormDataContent();
            foreach (var keyValuePair in nameValueCollection)
            {
                if (keyValuePair.Value != null)
                {
                    content.Add(new StringContent(keyValuePair.Value), keyValuePair.Key);
                }
                else
                {
                    content.Add(new StringContent(""), keyValuePair.Key);
                }
            }
            return content;
        }
        private FormUrlEncodedContent TokenRequestContent()
        {
            var parameters = new Dictionary<string, string>
                {
                    {"username", ""},
                    {"password", ""},
                    {"client", "referer " },
                    {"expiration", "1440" },
                    {"f", "pjson" },
                };
            var content = new FormUrlEncodedContent(parameters);
            return content;
        }
        public async Task<string> DownloadFileFromUrl(string outputFulName, string url, string token)
        {
            try
            {
                url = url + "?token=" + token;
                using (HttpClient http = new HttpClient())
                {
                    HttpResponseMessage response = http.GetAsync(url).Result;
                    if (response.IsSuccessStatusCode)
                    {
                       
                        try
                        {
                            var html = await response.Content.ReadAsStringAsync();

                            if (html.Contains("Invalid token")) return null;
                        }
                        catch
                        {

                        }
                        
                        
                        byte[] zipFileData = await response.Content.ReadAsByteArrayAsync();
                        // Save the zip file to disk
                        System.IO.File.WriteAllBytes(outputFulName, zipFileData);
                        return outputFulName;
                    }
                    else
                    {
                        MessageBox.Show(string.Format("{0}: {1} ({2})", url, (int)response.StatusCode, response.ReasonPhrase));
                    }
                }
            }
            catch (Exception exception)
            {
                return "";
            }
            return "";
        }

    }
}
