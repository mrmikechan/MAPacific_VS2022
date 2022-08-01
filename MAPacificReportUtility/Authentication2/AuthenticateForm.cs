using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GraphAuthentication
{
    //TODO: Add logic to disable close button until the authentication process has finished. Need to parse browser content and look for the word "close"

    public partial class AuthenticateForm : Form
    {
        //  private string message = "To sign in, use a web browser to open the page https://microsoft.com/devicelogin and enter the code FHPDNFLVD to authenticate.";
        //  private string info;

        private string message;
        public AuthenticateForm()
        {
            InitializeComponent();
            message = "Please copy and paste the device code into the page for verification: ";
        }

        public void NavigatePage(string info)
        {
            textBox1.Text =  message +ParseInfo(info) ;
            try
            {
                webBrowser1.Navigate(new Uri("https://microsoft.com/devicelogin"));
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        /// <summary>
        /// ParseInfo method is to parse out the device code from the response message sent from host.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private string ParseInfo(string text)
        {
            //need to add 4 to the first index to exclude code in the return string.
            int firstIndex = text.IndexOf("code") + 4;
            int lastIndex = text.LastIndexOf("to");
            string substring = text.Substring(firstIndex, lastIndex - firstIndex);
            return substring.Trim();
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (e.Url.AbsolutePath != (sender as WebBrowser).Url.AbsolutePath)
                return;
        }
    }
}
