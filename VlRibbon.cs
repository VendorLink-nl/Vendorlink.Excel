using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json.Linq;

namespace Vendorlink.Excel
{
    public partial class VlRibbon
    {
        private void VlRibbon_Load(object sender, RibbonUIEventArgs e)
        {


        }

        private void RefreshListButton_Click(object sender, RibbonControlEventArgs e)
        {
            QueryDropwdown.Items.Clear();

            var jToken = GetJsonFromRequest("/API/Queries/List", "GET", "");
            var jArray = (JArray)jToken;

            RibbonDropDownItem emptyItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            emptyItem.Label = "";
            emptyItem.Tag = null;
            QueryDropwdown.Items.Add(emptyItem);
            foreach (JToken jObject in jArray)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = jObject["Name"].Value<string>();
                item.Tag = jObject["Id"].Value<string>();
                QueryDropwdown.Items.Add(item);
            }
        }

        private JToken GetJsonFromRequest(string requestPath, string httpMethod, string body)
        {
            var accessToken = GetAccessToken();

            var webReq = WebRequest.Create($"http://{Properties.Settings.Default.Hostname}{requestPath}");
            webReq.Method = httpMethod;
            webReq.ContentType = "application/json";
            webReq.Headers.Add("Authorization", $"Bearer {accessToken}");

            if (!String.IsNullOrEmpty(body))
            {
                ASCIIEncoding encoding = new ASCIIEncoding();
                byte[] byte1 = encoding.GetBytes(body);
                // Set the content length of the string being posted.
                webReq.ContentLength = byte1.Length;
                // get the request stream
                Stream newStream = webReq.GetRequestStream();
                // write the content to the stream
                newStream.Write(byte1, 0, byte1.Length);
            }

            var webResp = (HttpWebResponse)webReq.GetResponse();

            if (webResp.StatusCode == HttpStatusCode.Unauthorized)
            {
                // we need a new token
                RenewToken();
                accessToken = GetAccessToken();
                // execute the request again, with the new token
                webReq.Headers["Authorization"] = accessToken;
                webResp = (HttpWebResponse)webReq.GetResponse();
                // if still unauthorized: login again
                if (webResp.StatusCode == HttpStatusCode.Unauthorized)
                {
                    webReq.Headers["Authorization"] = GetAccessToken(true);
                }
                // execute the request again, with the new token
                webResp = (HttpWebResponse)webReq.GetResponse();
            }

            var readStream = new StreamReader(webResp.GetResponseStream(), Encoding.UTF8);

            var jToken = ParseJson(readStream.ReadToEnd());
            return jToken;
        }

        private string GetAccessToken(bool forceLogin = false)
        {
            if (string.IsNullOrEmpty(Properties.Settings.Default.AccessToken))
            {
                var frm = new LoginForm();
                frm.ShowDialog();
            }

            return Properties.Settings.Default.AccessToken;
        }

        private void RenewToken()
        {
            var accessToken = GetAccessToken();
            var webReq = WebRequest.Create($"http://{Properties.Settings.Default.Hostname}/API/auth/extendtoken");
            webReq.Method = "POST";
            webReq.ContentType = "application/json";
            webReq.Headers.Add("Authorization", $"Bearer {accessToken}");

            webReq.ContentType = "application/json; charset=utf-8";
            string postData = $"{{\"RenewalToken\":\"{Properties.Settings.Default.RenewalToken}\"}}";
            ASCIIEncoding encoding = new ASCIIEncoding();
            byte[] byte1 = encoding.GetBytes(postData);
            // Set the content length of the string being posted.
            webReq.ContentLength = byte1.Length;
            // get the request stream
            Stream newStream = webReq.GetRequestStream();
            // write the content to the stream
            newStream.Write(byte1, 0, byte1.Length);
            // execute the request
            var webResp = webReq.GetResponse();

            var readStream = new StreamReader(webResp.GetResponseStream(), Encoding.UTF8);
            var jToken = (JObject)JToken.Parse(readStream.ReadToEnd());
            if (jToken != null && jToken["accessToken"] != null)
            {
                Properties.Settings.Default.AccessToken = jToken["accessToken"].Value<string>();
                Properties.Settings.Default.Save();
            }
        }


        private void Login()
        {

        }

        private JToken ParseJson(string json)
        {
            var retval = JToken.Parse(json);

            return retval;
        }

        private void BtnFillSheet_Click(object sender, RibbonControlEventArgs e)
        {
            if (QueryDropwdown.SelectedItem.Tag != null)
            {
                var id = int.Parse(QueryDropwdown.SelectedItem.Tag.ToString());
                var jToken = GetJsonFromRequest($"/API/Queries/Execute?id={id}", "GET", "");
                // we need the value of the first property: that contains the array
                var jArray = (JArray)jToken[QueryDropwdown.SelectedItem.Label];

                Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;

                var row = 1;
                var col = 1;
                var propNames = new List<string>();
                foreach (JProperty property in ((JObject)jArray[0]).Properties())
                {
                    propNames.Add(property.Name);
                    ((Range)sheet.Cells[row, col]).Value2 = property.Name;
                    col++;
                }

                foreach (JObject jObject in jArray)
                {
                    col = 1;
                    row++;
                    foreach (JProperty property in jObject.Properties())
                    {
                        var propVal = property.Value.ToString();
                        ((Range)sheet.Cells[row, col]).Value2 = propVal;
                        col++;
                    }
                }
            }

        }

        private void LoginButton_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.AccessToken = "";
            var frm = new LoginForm();
            frm.ShowDialog();
        }
    }
}
