using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using Microsoft.Identity.Client;

namespace VEYMDataParser
{
    public partial class MainForm : Form
    {
        private AuthenticationResult authentication;
        VEYMDataObjectManager ourGiantCollection;
        VEYMDataObjectManagerBETA ourGiantCollectionBETA;

        public MainForm(AuthenticationResult auth)
        {
            InitializeComponent();

            authentication = auth;
        }

        private void BtnScrape_Click(object sender, EventArgs e)
        {
            ScrapeHelper();
        }

        private void buttonBETA_Click(object sender, EventArgs e)
        {
            ScrapeHelperBETA();
        }

        private async void ScrapeHelper()
        {
            if (authentication != null)
            {
                string graphAPIEndpoint = "https://graph.microsoft.com/v1.0/users";
                string json;
                List<AllUsersDataObject.RootObject> allPages = new List<AllUsersDataObject.RootObject>();

                while (!String.IsNullOrEmpty(graphAPIEndpoint))
                {
                    json = await GetHttpContentWithToken(graphAPIEndpoint, authentication.AccessToken);
                    AllUsersDataObject.RootObject page = JsonConvert.DeserializeObject<AllUsersDataObject.RootObject>(json);
                    allPages.Add(page);
                    graphAPIEndpoint = page.nextLink;
                }

                ourGiantCollection = new VEYMDataObjectManager(allPages);
            }
        }

        private async void ScrapeHelperBETA()
        {
            if (authentication != null)
            {
                string graphAPIEndpoint = "https://graph.microsoft.com/beta/users";
                string json;
                List<AllUsersDataObjectBETA.RootObject> allPages = new List<AllUsersDataObjectBETA.RootObject>();

                while (!String.IsNullOrEmpty(graphAPIEndpoint))
                {
                    json = await GetHttpContentWithToken(graphAPIEndpoint, authentication.AccessToken);
                    AllUsersDataObjectBETA.RootObject page = JsonConvert.DeserializeObject<AllUsersDataObjectBETA.RootObject>(json);
                    allPages.Add(page);
                    graphAPIEndpoint = page.nextLink;
                }

                ourGiantCollectionBETA = new VEYMDataObjectManagerBETA(allPages);
            }
        }




        /// <summary>
        /// Perform an HTTP GET request to a URL using an HTTP Authorization header
        /// </summary>
        /// <param name="url">The URL</param>
        /// <param name="token">The token</param>
        /// <returns>String containing the results of the GET operation</returns>
        public async Task<string> GetHttpContentWithToken(string url, string token)
        {
            var httpClient = new System.Net.Http.HttpClient();
            System.Net.Http.HttpResponseMessage response;
            try
            {
                var request = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, url);
                //Add the token in Authorization header
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                response = await httpClient.SendAsync(request);
                var content = await response.Content.ReadAsStringAsync();
                return content;
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }

        private async void BtnMyself_Click(object sender, EventArgs e)
        {
            if (authentication != null)
            {
                string graphAPIEndpoint = "https://graph.microsoft.com/v1.0/me";
                string json;

                while (!String.IsNullOrEmpty(graphAPIEndpoint))
                {
                    json = await GetHttpContentWithToken(graphAPIEndpoint, authentication.AccessToken);
                }
            }
        }

        private async void BtnLogOff_Click(object sender, EventArgs e)
        {
            var accounts = await App.PublicClientApp.GetAccountsAsync();
            if (accounts.Any())
            {
                try
                {
                    await App.PublicClientApp.RemoveAsync(accounts.FirstOrDefault());
                    MessageBox.Show("You have logged off successfully!");
                }
                catch (MsalException ex)
                {
                    MessageBox.Show("Uh oh, error in Log Out");
                }
            }
        }
    }
}
