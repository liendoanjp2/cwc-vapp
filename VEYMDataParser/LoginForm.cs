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

namespace VEYMDataParser
{
    public partial class LoginForm : Form
    {
        public LoginForm()
        {
            InitializeComponent();
        }

        string[] scopes = new string[] { "user.read.all" };
                    

        private async void btnLogin_Click(object sender, EventArgs e)
        {
            AuthenticationResult credentials = await GetCredentials();
            MainForm theMainForm = new MainForm(credentials);
            theMainForm.Show();

            this.Hide();
        }

        private async Task<AuthenticationResult> GetCredentials()
        {
            AuthenticationResult authResult = null;
            var app = App.PublicClientApp;
            string tokenInfoText = string.Empty;

            var accounts = await app.GetAccountsAsync();
            var firstAccount = accounts.FirstOrDefault();

            try
            {
                authResult = await app.AcquireTokenSilent(scopes, firstAccount)
                    .ExecuteAsync();
            }
            catch (MsalUiRequiredException ex)
            {
                // A MsalUiRequiredException happened on AcquireTokenSilent. 
                // This indicates you need to call AcquireTokenInteractive to acquire a token
                System.Diagnostics.Debug.WriteLine($"MsalUiRequiredException: {ex.Message}");

                try
                {
                    authResult = await app.AcquireTokenInteractive(scopes)
                        .WithAccount(accounts.FirstOrDefault())
                        .WithPrompt(Prompt.SelectAccount)
                        .ExecuteAsync();
                }
                catch (MsalException msalex)
                {
                    MessageBox.Show($"Error Acquiring Token:{System.Environment.NewLine}{msalex}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error Acquiring Token Silently:{System.Environment.NewLine}{ex}");
            }
                return authResult;
        }
    }
}
