using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json.Linq;

namespace Vendorlink.Excel
{
    public partial class LoginForm : Form
    {
        public LoginForm()
        {
            InitializeComponent();
        }

        private void BtnLogin_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Hostname = txtHostname.Text;
            var webReq = WebRequest.Create($"http://{Properties.Settings.Default.Hostname}/API/auth/login");
            webReq.Method = "POST";
            webReq.ContentType = "application/json; charset=utf-8";
            string postData = $"{{\"u\":\"{txtUsername.Text}\",\"p\":\"{txtPassword.Text}\"}}";
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

            if (webResp.ContentLength <= 0)
            {
                MessageBox.Show("Login failed");
                return;
            }

            var readStream = new StreamReader(webResp.GetResponseStream(), Encoding.UTF8);
            var jToken =(JObject) JToken.Parse(readStream.ReadToEnd());
            if (jToken["displayName"] != null)
            {
                Properties.Settings.Default.DisplayName = jToken["displayName"].Value<string>();
                Properties.Settings.Default.UserId = jToken["userId"].Value<int>();
                Properties.Settings.Default.AccessToken = jToken["accessToken"].Value<string>();
                Properties.Settings.Default.RenewalToken = jToken["renewalToken"].Value<string>();
                Properties.Settings.Default.Save();
            }
            this.Close();
        }
    }
}
