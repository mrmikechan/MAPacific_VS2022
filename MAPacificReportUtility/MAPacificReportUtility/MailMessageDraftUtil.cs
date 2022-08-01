using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using MAPacificReportUtility.forms;
using Microsoft.Graph;
using GraphAuthentication;

namespace MAPacificReportUtility
{
    /// <summary>
    /// 7/2/2022
    /// MailMessageDraft used to use EWS Manage API with Microsoft.Exchange.WebServices.dll to create email drafts and attachments.
    /// However the Basic Authentication that comes with EWS does not work with the more modern Azure Active Directory with cloud accounts
    /// that uses MFA. There was a need to migrate over to Microsoft.Graph API to perform the email draft functionality.  Microsoft.Graph has
    /// quite a bit of requirements.  You will need to use Nuget Manager for package solution to install multiple packages:
    /// 
    /// Azure.Core.1.24.0
    /// Azure.Identity.1.6.0
    /// Microsoft.Bcl.AsyncInterfaces.6.0.0
    /// Microsoft.Extensions.Configuration.6.0.1
    /// Microsoft.Extensions.Configuration.Abstractions.6.0.0
    /// Microsoft.Extensions.Configuration.Binder.6.0.0
    /// Microsoft.Extensions.Configuration.FileExtensions.6.0.0
    /// Microsoft.Extensions.Configuration.Json.6.0.0
    /// Microsoft.Extensions.FileProviders.Abstractions.6.0.0
    /// Microsoft.Extensions.FileProviders.Physical.6.0.0
    /// Microsoft.Extensions.FileSystemGlobbing.6.0.0
    /// Microsoft.Extensions.Primitives.6.0.0
    /// Microsoft.Graph.4.31.0
    /// Microsoft.Graph.Core.2.0.9
    /// Microsoft.Identity.Client.4.44.0
    /// Microsoft.Identity.Client.Extensions.Msal.2.19.3
    /// Microsoft.IdentityModel.Abstractions.6.19.0
    /// Microsoft.IdentityModel.JsonWebTokens.6.19.0
    /// Microsoft.IdentityModel.Logging.6.19.0
    /// Microsoft.IdentityModel.Protocols.6.19.0
    /// Microsoft.IdentityModel.Protocols.OpenIdConnect.6.19.0
    /// Microsoft.IdentityModel.Tokens.6.19.0
    /// System.Buffers.4.5.1
    /// System.Diagnostics.DiagnosticSource.4.7.1
    /// System.IdentityModel.Tokens.Jwt.6.19.0
    /// System.IO.4.3.0
    /// System.Memory.4.5.4
    /// System.Memory.Data.1.0.2
    /// System.Net.Http.4.3.4
    /// System.Net.Http.WinHttpHandler.6.0.1
    /// System.Numerics.Vectors.4.5.0
    /// System.Runtime.4.3.0
    /// System.Runtime.CompilerServices.Unsafe.6.0.0
    /// System.Runtime.InteropServices.RuntimeInformation.4.3.0
    /// System.Security.Cryptography.Algorithms.4.3.0
    /// System.Security.Cryptography.Algorithms.4.3.1
    /// System.Security.Cryptography.Encoding.4.3.0
    /// System.Security.Cryptography.Primitives.4.3.0
    /// System.Security.Cryptography.ProtectedData.4.7.0
    /// System.Security.Cryptography.X509Certificates.4.3.0
    /// System.Text.Encodings.Web.6.0.0
    /// System.Text.Json.6.0.4
    /// System.Threading.Tasks.Extensions.4.5.4
    /// System.ValueTuple.4.5.0
    /// 
    /// MailMessageDraft uses the EWS Managed API assembly file, Microsoft.Exchange.WebServices.dll to handle the
    /// interfacing logic to Exchange Server 2007 SP1 or higher.
    /// Requirements for using EWS Managed API assembly is that we need to use .NET Framework 4.
    /// </summary>
    class MailMessageDraftUtil
    {
        private Dictionary<string, EmailDraft> _draftList;

        IProgress<int> myProgress;
        
        //Call GraphAPI via our graph service
        private GraphHelper graphHelper;

        public MailMessageDraftUtil()
        {
            _draftList = new Dictionary<string, EmailDraft>();
            graphHelper = null;
            myProgress = null;
        }

        /// <summary>
        /// CreateDraftMail method creates new EmailDraft and adds it to the _draftList dictionary. 
        /// </summary>
        /// <param name="To">target to send the email</param>
        /// <param name="Subject">email subject</param>
        /// <param name="Body">email body</param>
        /// <param name="attachment">file attachment</param>
        public void CreateDraftEmail(String To, String Subject, String Body, String attachment)
        {
            //check to see if draft email exists...
            if (_draftList.ContainsKey(To))
            {
                //it exists so we need to update the object and then insert it back into the dictionary.
                EmailDraft msg = _draftList[To];
                msg.Subject += Subject.Substring(Subject.IndexOf(":") + 1);
                msg.Attachments.Add(attachment);
                _draftList[To] = msg;
            }
            else
            {
                EmailDraft message = new EmailDraft();
                message.To = To;
                message.Subject = Subject;
                message.Body = "Have a great day!" + System.Environment.NewLine + System.Environment.NewLine;
                message.Body += "Accounting and Settlements";
                message.Attachments.Add(attachment);
                _draftList.Add(To, message);
            }

        }

        /// <summary>
        /// SaveDraftEmail runs through the dictionary and saves all email drafts into the user's Outlook draft folder.
        /// </summary>
        public void SaveDraftEmail(GraphAuthentication.GraphHelper helper)
        {
            //3-27-22 MAP reported issue in which the Email Draft feature was not working anymore. It would appear that changes were made on Office365/Outlook365 requiring clients to use TLS12 to communicate with Exchange WebService.
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

            graphHelper = helper;
            FloatingProgressBar progressBar = new FloatingProgressBar("Creating Emails and saving into Outlook Draft folder", EmailListSize()); ;
            myProgress = new Progress<int>(x => progressBar.UpdateProgressBar(x));
            progressBar.Show();

            //Use Tasks to create the draft emails in async way.
            System.Threading.Tasks.Task t = new System.Threading.Tasks.Task(() => CreateDraftEmailAsync(progressBar));
            t.Start();
            t.Wait();
        }

        private async void CreateDraftEmailAsync(FloatingProgressBar mBar)
        {
            //converted Exchange.Web.Services related logic over to Microsoft.Graph.
            Microsoft.Graph.Message msg;
            if (_draftList.Count > 0)
            {
                //count used for the progress panel
                int count = 0;
                foreach (string keys in _draftList.Keys)
                {
                    try
                    {
                        count++;
                        //    if (updateProgress != null)
                        if (myProgress != null)
                        {
                            mBar.UpdateProgressBar(count);
                           // myProgress.Report(count);
                           // progressBar.PercentageComplete = (count / _draftList.Count) * 100;
                        }
                        msg = _draftList[keys].ConvertToEmailMessage(graphHelper);
                        
                        await graphHelper.getClient().Me.Messages.Request().AddAsync(msg);

                    }
                    catch (Exception e)
                    {
                        if (_draftList[keys].To.Length == 0)
                        {
                            System.Windows.Forms.MessageBox.Show("Unable to Save Email Draft because there is no target email address: " + _draftList[keys].Subject, "Email Draft Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                        }
                        else
                        {
                            System.Windows.Forms.MessageBox.Show("Exception encountered in saving draft email: " + e.Message, "Email Draft Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                        }
                    }
                }
                //Finished with creating draft email messages. Close the progressbar winform.
                mBar.BeginInvoke(new Action(() => mBar.Close()));
            }
        }

        /// <summary>
        /// ResetList empties out the Dictionary containing the Email messages
        /// </summary>
        public void ResetList()
        {
            _draftList.Clear();
        }

        /// <summary>
        /// isListEmpty returns true or false if there is any content in the dictionary list.
        /// </summary>
        /// <returns></returns>
        public bool isListEmpty()
        {
            if (_draftList.Count <= 0)
                return true;
            else
                return false;
        }

        /// <summary>
        /// EmailListSize returns the number of email drafts in the Dictionary.
        /// </summary>
        /// <returns></returns>
        public int EmailListSize()
        {
            return _draftList.Count;
        }
    }
    /// <summary>
    /// Class file to encapsulate data contained in a email.
    /// </summary>
    public class EmailDraft
    {
        public EmailDraft()
        {
            _attachments = new List<string>();
        }

        private string _to = "";
        public string To
        {
            get { return _to; }
            set { _to = value; }
        }

        private string _subject = "";
        public string Subject
        {
            get { return _subject; }
            set { _subject = value; }
        }

        private string _body = "";
        public string Body
        {
            get { return _body; }
            set { _body = value; }
        }

        private List<string> _attachments;
        public List<string> Attachments
        {
            get { return _attachments; }
            set { _attachments = value; }
        }

        /// <summary>
        /// ConvertToEmailMessage function uses Microsoft.Graph to create a email draft and attach the pdf file to the message.
        /// </summary>
        /// <param name="graphHelper">Class that encapsulates the Microsoft.Graph function calls</param>
        /// <returns>Microsoft.Graph.Message</returns>
        /// <exception cref="System.NullReferenceException"></exception>
        public Message ConvertToEmailMessage(GraphHelper graphHelper)
        {
            if (graphHelper == null)
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            GraphServiceClient userClient = graphHelper.getClient();

            Microsoft.Graph.Message message = new Message();
            //VS6159 Fix Draft meail address issue. Sometines the address generated is not recognized as a valid email address by Outlook.  User has to retype the
            //email address for it to be recognized.

            //If email address is not a real address but a contact group name...
            if (To.Contains("@")) //contact group name that is like a email address is the correct format for group name /distribution list.
            {
                message.ToRecipients = new Recipient[]
                {
                    new Recipient
                    {
                        EmailAddress = new Microsoft.Graph.EmailAddress
                        {
                            Address = To
                        }
                    }
                };
            }
            else
            {
                //Unfortunately EWS for some odd reason is not able to recognize the email contact group name. It only thinks of it as a string and not an actual contact group name
                //and therefore no address is associated with the group name placed into the email's To value. As a work around we have to find the contact group and then logically
                //add each member's email address to the draft.
                //       ItemView view = new ItemView(9999);
                //       view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ContactGroupSchema.DisplayName, ContactGroupSchema.DisplayTo);
                //       FindItemsResults<Item> contactGroupItems = service.FindItems(WellKnownFolderName.Contacts, To, view);

                //MS Graph API to get all the Groups
                //      IGraphServiceGroupsCollectionPage groupContactPage = graphHelper.GetContactGroups();
    /* 
                var groupContactPage = graphHelper.GetContactGroups();
                var foundGroup = from contact in groupContactPage where (contact.DisplayName == To) select contact;
                if (foundGroup != null)
                {
                    List<Recipient> recipientsList = new List<Recipient>();
                    var group = foundGroup.ElementAt(0);
                    var members = graphHelper.GetGroupMembers(group.Id);
                    Recipient recipient;
                    foreach (User member in members)
                    {
                        if (member.Mail != null)
                        {
                            recipient = new Recipient();
                            recipient.EmailAddress.Address = member.Mail;
                            recipientsList.Add(recipient);
                        }
                    }
                    // var recipient = new Recipient();
                    // recipient.EmailAddress = member.Mail;

                    message.ToRecipients = recipientsList;
                    // message.ToRecipients = "ms@f.com";
                }
*/
//                else
 //               {

                    message.ToRecipients = new Recipient[]
                    {
                        new Recipient
                        {
                            EmailAddress = new Microsoft.Graph.EmailAddress
                            {
                                Address = To
                            }
                        }
                    };

 //               }

                message.Subject = Subject;
                message.Body = new ItemBody { ContentType = Microsoft.Graph.BodyType.Text };
                message.Body.Content = Body;

                MessageAttachmentsCollectionPage attachments = new MessageAttachmentsCollectionPage();
                string contentType = "";
                foreach (string attch in Attachments)
                {
                    if (attch.EndsWith("pdf"))
                    {

                        contentType = "application/pdf";
                    }
                    


                        Microsoft.Graph.FileAttachment attachmentItem = new Microsoft.Graph.FileAttachment()
                        {
                            Name = System.IO.Path.GetFileName(attch),
                            ContentBytes = System.IO.File.ReadAllBytes(attch),
                            ContentType = contentType
                        };
                    attachments.Add(attachmentItem);
                }

                message.Attachments = attachments;

            }
            return message;
        }
    }
}
