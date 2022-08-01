using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Text;
using OvDotNet;

/*
 * Change   Date    Name    Details
 * -------------------------------------------------------------------------------------------------------------------------------------------
 * $d01     033012  MC      OVdotNetApi is problematic under 64bit OS. Specific calls such as ovApi.Text and  ovApi.CrtGet(5, 0, ovApi.Columns)
 *                          returns odd data or empty data where there should be real data. I think there are some threading issues that is causing
 *                          this intermittent problem.  As a work around I have changed some logic so that we process or retrieve the data from
 *                          the ovApi in a manner that we are guaranteed to have data.
 */


namespace MAPacificReportUtility
{
    class ProcessVisaDPSReport : ProcessReport
    {
        private bool debug = false; //flag used to generate trace outputs...

        public ProcessVisaDPSReport(OvDotNetApi inApi)
        {
            _branchInfo = new ClientBranch();
            _sarpage = 0;
            subPage = 1;
            _isglobalsummary = false;
            ovApi = inApi;
        }


        #region properties


        private OvDotNetApi ovApi;
        private string apiText = ""; //$d01 optimization used to hold screen content from ovApi...

        //VS4142 Modification to work with Changes to VisaDPS Report
        //Contents of reports are changing and being split into seperate GPR and GIFT reports with their own BIN
        //number.
        //GPR BIN#:  99586001
        //GIFT BIN#: 99586002
        const string MAP_GPR_BIN = "99586001";
        const string MAP_GIFT_BIN = "99586002";
        //VS4583 Enhance tool to work with a new entity named Beken which handles private bank reports.
        const string BEKEN_GPR_BIN =  "99818001";
        //There isn't one for Beken GIFT BIN yet...so we set arbitrary value of 0.
        const string BEKEN_GIFT_BIN = "00000000";

        //VS4142
        //Because we can run seperate reports now, the logic to check for report date may not work all the time.
        //Use a flag instead.
        bool isDateChecked = false;

        public struct ClientBranch_Category
        {
            static public string GPR_DETAIL
            {
                get { return "GPR Detail"; }
            }

            static public string GPR_SUMMARY
            {
                get { return "GPR Summary"; }
            }

            static public string GIFT_DETAIL
            {
                get { return "Gift Detail"; }
            }

            static public string GIFT_SUMMARY
            {
                get { return "Gift Summary"; }
            }

            static public string IGNORE
            {
                get { return "Ignore"; }
            }

            static public string PREPAID_TOTAL_SUMMARY
            {
                get { return "Prepaid Total Summary"; }
            }
        };

        private struct ClientBranchBin
        {
            public const int PrePaidCardsTotal = 0;
            public const int GPRSummary = 1;
            public const int GiftCardSummary = 2;
            public const int GPRBranchDetail = 3;
            public const int GiftBranchDetail = 4;
            public const int IgnoreReport = 5;
            public const int UmbrellaReport = 6;
            public const int UnrecognizedID = 7;
        }

        private ClientBranch _branchInfo;
        public ClientBranch BranchInfo
        {
            get { return _branchInfo; }
        }

        private Boolean sameClient = false;
        private int subPage = 0;

        private bool _reportfinish = true; //start off with true because the report isn't running when application is launched.
        public bool ReportFinish
        {
            get { return _reportfinish; }
            set
            {
                _reportfinish = value;
            }
        }

        public bool _reportcancel = false;
        public bool ReportCancel
        {
            get { return _reportcancel; }
            set
            {
                _reportcancel = value;
            }
        }

        /// <summary>
        /// isPreviousSarPage boolean value is used to determine if someone press the function key to go to the previous
        /// screen of the report. Note that this is an edge case because in a perfect world this automated parser will only issue
        /// next page event to get the next screen and not go backwards! However if the user decides to hit the function key on
        /// OV to go back a page then we need to adjust for that.
        /// </summary>
        private bool isPreviousSarPage = false;

        private int _sarpage;
        /// <summary>
        /// </summary>
        private int SarPage
        {
            get { return _sarpage; }
            set
            {
                _sarpage = value;
                BranchChanged(this.BranchInfo);
                sameClient = false;
            }
        }

        private bool _isglobalsummary;
        /// <summary>
        /// IsGlobalSummary is used to determine if we have reached the end of the report where the global summary resides.
        /// </summary>
        public bool IsGlobalSummary
        {
            get { return _isglobalsummary; }
        }

        private string _reportType;
        /// <summary>
        /// Property to determine what report type that we are parsing. Basically this value is based
        /// on the BIN numper parsed from the report. ReportType can be the following value: MAP Prepaid, Map Gift, Beken Prepaid, or Beken Gift.
        /// </summary>
        public String ReportType
        {
            get { return _reportType; }
            set
            {
                if (value != null)
                {
                    _reportType = value;
                }
            }
        }

        #endregion

        #region parse logic
        /// <summary>
        /// ProcessData parses the data text from the OvDotNetApi.
        /// </summary>
        public void ParseData()
        {
            //$d01 use a loop to retrieve the OV screen text. The OV screen will always have some sort of texts. Its just that the ovApi
            //has issues retrieveing that content when the host is writing to the screen and we are trying to retrieve it at the same time....
            int counter = 0;
            do
            {
                //$d01
                //More strange behaviors on 64bit OS platform.  Sometimes the ovApi.Text returns blank value when there is data on the OV screen.
                //As a work around, we will try to get the ovApi.Text untill we actually have data in there...
                apiText = ovApi.Text;
                if(debug)
                    System.Diagnostics.Trace.WriteLine("ParseData datalength: " + apiText.Length + " try number: " + counter++ + "\r\n" +  apiText);
            }while(apiText.Length <= 3);
            
            //
            //check for a valid page to parse. A valid page would contain the SARPAGE keyword.
            //What we want is the "SARPAGE 1" keyword and the digit accompanied with it.
            if (isTargetTextPresent(apiText, new Regex("(SARPAGE\\s\\d*)")))
            {
                //lets get the page number for SARPAGE which will help us in determining a new sub client ID is encountered or not.
                SetSarPage(apiText);

                //Find the BIN report value so that we can determine if this report is GPR or GIFT report
                SetBIN(apiText);

                //lets find the SubClientID now. The report page that corresponds to the row which contains SubClientID and Name
                //always start with "DD5030 -D01".  We can look for this token and then search for the SubClientID and Name respectively.
                //Also since we are on a screen that is a valid report page, we know that the row containing the SubClientID and Name is
                //always the same row. So we can just retrieve the contents from that row instead of having to use the whole screen content.
                try
                {
                    //SetSubClientIDandName(ovApi.CrtGet(5, 0, ovApi.Columns));
                    //$d01 Instead of calling the CrtGet which can be flaky in 64bit OS... we just get the Text that was retrieved earlier.
                    SetSubClientIDandName(apiText);
                }
                catch (Exception e1)
                {
                    //working around strange behavior with the ovApi.CrtGet cmd.  This command does not always return the correct content. For example ovApi.Text will
                    //contains the OV screent content. Subsequently using ovApi.CrtGet to get content on row 5, column 0 sometimes returns no data.
                    string text = ovApi.CrtGet(5, 0, ovApi.Columns);
                    SetSubClientIDandName(text);
                }
                //Set the category value based on the SubClientID..

                try
                {

                    SetCategory();
                }
                catch (Exception e3)
                {
                    System.Diagnostics.Trace.WriteLine(e3.Message);
                }
                //Set the transaction type...
                try
                {
                    //SetTransactionType(ovApi.CrtGet(6, 0, (ovApi.Rows - 6) * ovApi.Columns));
                    //$d01 Instead of calling the CrtGet which can be flaky in 64bit OS... we just get the Text that was retrieved earlier.
                    SetTransactionType(apiText);
                }
                catch (Exception e2)
                {
                    string text = ovApi.CrtGet(6, 0, (ovApi.Rows - 6) * ovApi.Columns);
                    SetTransactionType(text);
                }
            }
        }

        //used only for getting the page being processed. 
        public int getSarPage()
        {
            return SarPage;
        }

        /// <summary>
        /// SetSarPage parses the text content and searches for the SARPAGE keyword and retrieves the integer value
        /// representing the page.
        /// </summary>
        /// <param name="inText">string which contains current OV session terminal screen text</param>
        public void SetSarPage(String inText)
        {
            Regex sarPage = new Regex("(SARPAGE\\s\\d*)");
            string txt = inText;
            MatchCollection mCollection = sarPage.Matches(txt);
            if (mCollection != null && mCollection.Count > 0)
            {
                txt = mCollection[0].Value;
                txt = txt.Substring(7); //7 is the length for the word SARPAGE;
                txt.Trim();

                int page = 0;
                try
                {
                    page = System.Convert.ToInt32(txt);

                    //VS4142 If we start GPR report it may not start at page 1 and thus
                    //the logic to check page == 1 fails.
                    if(!isDateChecked)
                    {
                        checkReportDate(inText);
                        isDateChecked = true;
                    }
                }
                catch (System.Exception ex)
                {
                    //do not update Sarpage..
                    return;
                }
                //only update Sarpage if we have higher page numbers...
                if (page > SarPage)
                {
                    SarPage = page;
                    //new client branch...
                    _branchInfo = new ClientBranch();
                    subPage = 1; //reset the subpage for the new client branch...
                    isPreviousSarPage = false;
                }
                else if (page < SarPage)//somehow we went to the previous screen...
                {
                    isPreviousSarPage = true;
                }
                else
                {
                    //still the same client ID but different page of report.
                    sameClient = true;
                    subPage++;

                    //if (_isglobalsummary && ReportFinish)//subPage >= 3) //We only need the data from first 2 pages of Global Summary.
                    //{
                    //   // _reportfinish = true;
                    //    EndofReport(this.Data);
                    //}
                }
            }
        }

        /// <summary>
        /// Set the BIN string value in the ClientBranch so that we know what report it falls under.
        /// </summary>
        /// <param name="inText"></param>
        public void SetBIN(string inText)
        {
            string text = inText;
            bool binSet = false;
            if (text.Length > 0)
            {
                Regex pattern = new Regex("-99586001");
                MatchCollection mCollection = pattern.Matches(text);

                //MAP GPR report
                if (mCollection.Count > 0)
                {
                    BranchInfo.BIN = MAP_GPR_BIN;
                    binSet = true;
                    ReportType = "MAP GPR";
                }

                pattern = new Regex("-99586002");
                mCollection = pattern.Matches(text);

                //MAP GIFT report
                if (mCollection.Count > 0)
                {
                    BranchInfo.BIN = MAP_GIFT_BIN;
                    binSet = true;
                    ReportType = "MAP GIFT";
                }

                //Beken GPR report
                pattern = new Regex("-99818001");
                mCollection = pattern.Matches(text);

                if (mCollection.Count > 0)
                {
                    BranchInfo.BIN = BEKEN_GPR_BIN;
                    binSet = true;
                    ReportType = "Beken GPR";
                }

                //Beken GIFT report -- if all else fails then we default to Beken Gift...
                if (!binSet)
                {
                    BranchInfo.BIN = "99586002";
                    ReportType = "Beken Gift";
                }
            }
        }

        public void SetSubClientIDandName(String inText)
        {
            string text = inText;
            string clientID = "";
            string clientName = "";
            MatchCollection mCollection;

            //If we have a new subpage page and it is still for the same Sub Client ID and Name, then we do not
            //need to parse that information again.
            if (sameClient || isPreviousSarPage)
            {
                return;
            }

            //find the starting location to begin looking for the clientID
            int strtIndex = text.IndexOf("-D01", 0);
            if (strtIndex > 0)
            {
                text = text.Substring(strtIndex + 4);
                text = text.Trim();
            }

            //There is no definitive way to determine the possible boundaries for the Sub Client ID value. Normally they would be
            //5 chars long and end with 3 digits. However it is possible that the Sub Client ID value could be all text and no digits. 
            //Such as "CALCOE FCU and BEEHIVE". The problem is if the value is all text how de we know how long the word is and if it is
            //more than 1 word long such as ABC DEF HIJ?
            Regex pattern = new Regex("(\\w*)");
            mCollection = pattern.Matches(text);

            //retrieve the SubclientID from the parsed results if it exists...
            if (mCollection.Count > 0 && !mCollection[0].Value.Trim().StartsWith("RUNTIME", StringComparison.CurrentCultureIgnoreCase))
            {
                clientID = mCollection[0].Value;

                //find the index for the first occurence of the clientID
                int index1 = text.IndexOf(clientID);

                int endIndex = text.LastIndexOf("RUNTIME");
                //optimized logic based on reports in which majority of the time the SubClientID values that is all alpha characters
                //seem to have the same value for their Name also. So if we can find the repeating value and its index position
                //then we can essentially figure out the SubClient ID Name and Name respectively!
                int index2 = text.LastIndexOf(clientID, StringComparison.CurrentCultureIgnoreCase);

                //now we need to determine if thi value is all alpha characters or alphanumeric characters. If it is
                //alphanumeric characters then we are essentially done. If not then we need to process some more to make sure
                //we got the whole Sub Client ID name..
                if (this.IsAlpha(clientID))
                {
                    //if the two values are different then we have multiple same word in the text...
                    if (index1 != index2)
                    {
                        clientID = text.Substring(index1, index2 - 2); //- 2 offset is to compensate for special chars such as * that my be part of the Name.     
                        clientName = text.Substring(index2 - 1, endIndex - (index2 - 1));
                    }
                    else //there is only one occurence of this word so more than likely that the Sub Client ID is not a multiple word value.
                    {
                        clientName = text.Substring(index1 + clientID.Length, endIndex - (index1 + clientID.Length));
                    }

                }
                else
                {
                    clientName = text.Substring(index1 + clientID.Length, endIndex - (index1 + clientID.Length));
                }

                if (clientName.Length > 0)
                {
                    BranchInfo.Name = clientName.Trim();
                }

                if (clientID.Length > 0)
                {
                    BranchInfo.ClientID = clientID.Trim();
                }
            }
            else
            {   //No ClientID value just empty strings upt to RUNTIME tag.
                //we have reached the global summary page! There are no values for the BranchInfo.. Instead we
                //want to accumulate the pages of text.
                _isglobalsummary = true;
            }



        }

        public void SetCategory()
        {
            //If we have a new subpage and it is still for the same Sub Client ID and Name, then we do not
            //need to parse that information again.
            if ((sameClient && subPage > 1) || isPreviousSarPage) //_isglobalsummary
            {
                return;
            }

            if (BranchInfo != null)
            {
                
                Regex pattern = new Regex("000$");
                //XXX000 -- Prepaid Cards Total Summary
                if (pattern.IsMatch(BranchInfo.ClientID))
                {
                    //There are Subclient IDs that have a valid ID but in the Name value they have content in there such as
                    //"Not in Use, or Test". We don't want to place them into IgnoreReport bin.
                    if (isVAlidIgnoreReport())
                        BranchInfo.Group = ClientBranchBin.IgnoreReport;
                    else
                        BranchInfo.Group = ClientBranchBin.PrePaidCardsTotal;
                    return;
                }

                //VS4142 Based on current Trace file from MAPacific, not sure if Visa is taking out the Z from ClientID names. What
                //we do know is that they will have a different BIN#, so we can track by that to determine GPR of GIFT report type.

                if (BranchInfo.BIN.Equals(MAP_GPR_BIN))
                {
                    //XXXX100 -- GPR Summary
                    //pattern = new Regex("^[^Zz][A-Za-z]+100$");
                    pattern = new Regex("^[A-Za-z]+100$");
                    if (pattern.IsMatch(BranchInfo.ClientID))
                    {
                        //There are Subclient IDs that have a valid ID but in the Name value they have content in there such as
                        //"Not in Use, or Test". We don't want to place them into IgnoreReport bin.
                        if (isVAlidIgnoreReport())
                            BranchInfo.Group = ClientBranchBin.IgnoreReport;
                        else
                            BranchInfo.Group = ClientBranchBin.GPRSummary;
                        return;
                    }

                    //XXX1NN -- GPR Branch Detail
                    //pattern = new Regex("^[^Zz][A-Za-z]+[0-9]+$");
                    pattern = new Regex("^[A-Za-z]+[0-9]+$");
                    if (pattern.IsMatch(BranchInfo.ClientID))
                    {
                        //There are Subclient IDs that have a valid ID but in the Name value they have content in there such as
                        //"Not in Use, or Test". We don't want to place them into IgnoreReport bin.
                        if (isVAlidIgnoreReport())
                            BranchInfo.Group = ClientBranchBin.IgnoreReport;
                        else
                            BranchInfo.Group = ClientBranchBin.GPRBranchDetail;
                        return;
                    }
                }

                if (BranchInfo.BIN.Equals(MAP_GIFT_BIN))
                {
                    //ZXXX100 -- Gift Card Summary
                    //pattern = new Regex("^[Zz][A-Za-z]+100$");
                    pattern = new Regex("^[A-Za-z]+100$");
                    if (pattern.IsMatch(BranchInfo.ClientID))
                    {
                        //There are Subclient IDs that have a valid ID but in the Name value they have content in there such as
                        //"Not in Use, or Test". We don't want to place them into IgnoreReport bin.
                        if (isVAlidIgnoreReport())
                            BranchInfo.Group = ClientBranchBin.IgnoreReport;
                        else
                            BranchInfo.Group = ClientBranchBin.GiftCardSummary;
                        return;
                    }



                    //ZXXX1NN -- Gift Branch Detail
                    //pattern = new Regex("^[Zz][A-Za-z]+[0-9]+$");
                    pattern = new Regex("[A-Za-z]+[0-9]+$");
                    if (pattern.IsMatch(BranchInfo.ClientID))
                    {
                        if (isVAlidIgnoreReport())
                            BranchInfo.Group = ClientBranchBin.IgnoreReport;
                        else
                            BranchInfo.Group = ClientBranchBin.GiftBranchDetail;
                        return;
                    }
                }

                if (isVAlidIgnoreReport())
                {
                    BranchInfo.Group = ClientBranchBin.IgnoreReport;
                    return;
                }

                //All characters -- Unrecognized Subclient ID.
                if (IsAlphaAndSpace(BranchInfo.ClientID))
                {
                    BranchInfo.Group = ClientBranchBin.UnrecognizedID;
                    return;
                }
            }
        }

        private Boolean isVAlidIgnoreReport()
        {
            Regex pattern;
            //All numbers, Name may contain "Test" or "Not in Use" -- Ignore
            if (IsNumber(BranchInfo.Name))
            {
                return true;
            }

            pattern = new Regex("Test|TEST");
            if (pattern.IsMatch(BranchInfo.Name))
            {
                return true;
            }

            pattern = new Regex("Not in Use|Not In Use|not in use|NOT IN USE");
            if (pattern.IsMatch(BranchInfo.Name))
            {
                return true;
            }
            return false;
        }

        public void SetTransactionType(String inText)
        {
            #region global summary logic handling
            if (IsGlobalSummary)
            {
                int index = 0;
              //  int endindex = 0;
                //if (subPage < 3)
                if(!ReportFinish)
                {
                    //logic to append the global summary report into the data object.
               //     if (subPage == 1) //1st page of global summary report
               //     {
                        //                       Data += "                            MAP Prepaid                            ";
                        //                       Data += Environment.NewLine;
                        //find second instance of "TRANSACTION TYPE" so that we know where to start getting the data...
                        index = inText.LastIndexOf("TRANSACTION TYPE");
                        Data += inText.Substring(index);
                        Data += Environment.NewLine;

                        //VS4080
                        if (inText.IndexOf("NET CHANGE") > 0)
                        {
                            ReportFinish = true;
                            EndofReport(this.Data);
                        }
                //    }

                //    if (subPage > 1)//== 2) //2nd page of global summary report
                //    {
                //        //                       Data += "--------------------------------------------------------------------------------";
                //        //                       Data += Environment.NewLine;

                //        try
                //        {
                //            //find index for "*TOTAL CARD ACTIVITY"
                //            index = inText.IndexOf("*TOTAL CARD ACTIVITY"); //starting position
                //            endindex = inText.IndexOf("FOREIGN ACTIVITY");  //ending position
                //            Data += inText.Substring(index, endindex - index);
                //        }
                //        catch (Exception ex)
                //        {
                //            //if we get into here that means there is no "FOREIGN ACTIVITY" string on page two because
                //            //there was too many other info and the Foreign Activity cut put into the third page.
                //            //What we know then is that the data runs to the end of the screen
                //            if (endindex < 0)
                //            {
                //                try
                //                {
                //                    //set the endindex to be inText.Length..
                //                    endindex = inText.Length;
                //                    Data += inText.Substring(index, endindex - index);
                //                }
                //                catch (Exception ex2)
                //                {
                //                    System.Windows.Forms.MessageBox.Show(ex2.Message, "Global Summary Parsing Exception", System.Windows.Forms.MessageBoxButtons.OK);
                //                }
                //            }
                //            else
                //                System.Windows.Forms.MessageBox.Show(ex.Message, "Global Summary Parsing Exception", System.Windows.Forms.MessageBoxButtons.OK);
                //        }
                //    }
                }
                return;
            }
            #endregion

            //If we have a new subpage and it is still for the same Sub Client ID and Name, then
            //we only need to parse subpages 1 and 2. No need to parse page 3 or higher.
            if (sameClient && subPage > 2 || isPreviousSarPage) //&& _isglobalsummary)
            {
                return;
            }

            bool haveData = false; //flag used to determine if we need to process the Total...
            MatchCollection mCollection = null;
            TransactionType aTransaction = null;

            try
            {
                //need to handle the different categories of transactions.
                //LOADS/FI FUNDS TRANSFER             1          115.00 CR        0.00              115.00 CR
                Regex pattern = new Regex("(\\sLOADS/FI FUNDS TRANSFER\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.LOADS_FI_FUNDS_TRANSFER;
                    parseTransactionString(foundValue, aTransaction);
                    haveData = true;
                }

                //RELOADS/FI FUNDS TRNSFER            1          115.00 CR        0.00              115.00 CR
                pattern = new Regex("(\\sRELOADS/FI FUNDS TRNSFER\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.RELOADS_FI_FUNDS_TRANSFER;
                    parseTransactionString(foundValue, aTransaction);
                    haveData = true;
                }

                //UNLOADS/FI FUNDS TRNSFER
                pattern = new Regex("(\\sUNLOADS/FI FUNDS TRNSFER\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.UNLOADS_FI_FUNDS_TRANSFER;
                    parseTransactionString(foundValue, aTransaction);
                    haveData = true;
                }

                //LOADS/MERCH POS FUNDING
                pattern = new Regex("(\\sLOADS/MERCH POS FUNDING\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.LOADS_MERCH_POS_FUNDING;
                    parseTransactionString(foundValue, aTransaction);
                    haveData = true;
                }

                //VS4637 Add LOADS_BYPASS per Orlando
                pattern = new Regex("(\\sLOADS/BY-PASS\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.LOADS_BYPASS;
                    parseTransactionString(foundValue, aTransaction);
                    haveData = true;
                }
                //RELOADS/BY-PASS
                pattern = new Regex("(\\sRELOADS/BY-PASS\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.RELOADS_BYPASS;
                    parseTransactionString(foundValue, aTransaction);
                    haveData = true;
                }

                //VS3594 Add UNLOADS/BY-PASS
                //UNLOADS/BY-PASS 
                pattern = new Regex("(\\sUNLOADS/BY-PASS\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.UNLOADS_BYPASS;
                    parseTransactionString(foundValue, aTransaction);
                    haveData = true;
                }

                //VS4637 Add LOADS_MERCHANT_ACQ per Orlando
                pattern = new Regex("(\\sLOADS/MERCHANT ACQUIRER\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.LOADS_MERCHANT_ACQ;
                    parseTransactionString(foundValue, aTransaction);
                    haveData = true;
                }

                //RELOADS/MERCHANT ACQ
                pattern = new Regex("(\\sRELOADS/MERCHANT ACQ\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.RELOADS_MERCHANT_ACQ;
                    parseTransactionString(foundValue, aTransaction);
                    haveData = true;
                }

                //UNLOADS/MERCHANT ACQ
                pattern = new Regex("(\\sUNLOADS/MERCHANT ACQ\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.UNLOADS_MERCHANT_ACQ;
                    parseTransactionString(foundValue, aTransaction);
                    haveData = true;
                }

                //*TOTAL LOAD/UNLOAD ACTIVITY
                if (haveData)
                {
                    pattern = new Regex("(\\s*TOTAL LOAD/UNLOAD ACTIVITY\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                    mCollection = pattern.Matches(inText);
                    if (mCollection.Count > 0)
                    {
                        string foundValue = mCollection[0].Value;
                        aTransaction = new TransactionType();
                        aTransaction.Transaction = TransactionType.TransactionOption.TOTAL_LOAD_UNLOAD_ACTIVITY;
                        parseTransactionString(foundValue, aTransaction);
                        haveData = true;
                    }
                }

                //MANUAL ADJUSTMENTS
                pattern = new Regex("(\\sMANUAL ADJUSTMENTS\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.MANUAL_ADJUSTMENT;
                    parseTransactionString(foundValue, aTransaction);
                }

                //VS4731 and VS4732 add logic to parse Card Activity transactions

                //Card Activity transactions
                //LOAD DISPUTES
                pattern = new Regex("(\\sLOAD DISPUTES\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.LOAD_DISPUTES;
                    parseTransactionString(foundValue, aTransaction);
                    haveData = true;
                }

                //PURCHASES/QUASI CASH
                pattern = new Regex("(\\sPURCHASES/QUASI CASH\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.PURCHASES_QUASI_CASH;
                    parseTransactionString(foundValue, aTransaction);
                    haveData = true;
                }

                //PURCHASES WITH CASH BACK
                pattern = new Regex("(\\sPURCHASES WITH CASH BACK\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.PURCHASES_WITH_CASH_BACK;
                    parseTransactionString(foundValue, aTransaction);
                    haveData = true;
                }

                //AFT - AA/PP
                pattern = new Regex("(\\sAFT - AA/PP\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.AFT_AA_PP;
                    parseTransactionString(foundValue, aTransaction);
                    haveData = true;
                }

                //PURCHASE RETURNS
                pattern = new Regex("(\\sPURCHASE RETURNS\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.PURCHASE_RETURNS;
                    parseTransactionString(foundValue, aTransaction);
                    haveData = true;
                }

                //MANUAL CASH
                pattern = new Regex("(\\sMANUAL CASH\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.MANUAL_CASH;
                    parseTransactionString(foundValue, aTransaction);
                    haveData = true;
                }

                //ATM CASH
                pattern = new Regex("(\\sATM CASH\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.ATM_CASH;
                    parseTransactionString(foundValue, aTransaction);
                    haveData = true;
                }

                //BALANCE INQUIRIES
                pattern = new Regex("(\\sBALANCE INQUIRIES\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.BALANCE_INQUIRIES;
                    parseTransactionString(foundValue, aTransaction);
                    haveData = true;
                }

                //REWARDS
                pattern = new Regex("(\\sREWARDS\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.REWARDS;
                    parseTransactionString(foundValue, aTransaction);
                    haveData = true;
                }

                //*TOTAL CARD ACTIVITY
                pattern = new Regex("(\\s*TOTAL CARD ACTIVITY\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.TOTAL_CARD_ACTIVITY;
                    parseTransactionString(foundValue, aTransaction);
                    haveData = true;
                }
                //VS4731 and VS4732 end

                //VS6160 Add ACH_DIRECT_DEPOSIT and RELOADS/MONEY_TSFR_RCVD
                pattern = new Regex("(\\sACH DIRECT DEPOSIT\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.ACH_DIRECT_DEPOSIT;
                    parseTransactionString(foundValue, aTransaction);
                    haveData = true;
                }

                pattern = new Regex("(\\sRELOADS/MONEY TSFR RCVD\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]*\\s*\\d*\\s*[0-9,.]+\\s*\\d*\\s*[0-9,.]*\\s+[A-Z]+)");
                mCollection = pattern.Matches(inText);
                if (mCollection.Count > 0)
                {
                    string foundValue = mCollection[0].Value;
                    aTransaction = new TransactionType();
                    aTransaction.Transaction = TransactionType.TransactionOption.RELOADS_MONEY_TRSFR_RCVD;
                    parseTransactionString(foundValue, aTransaction);
                    haveData = true;
                }

            }
            catch (Exception ex)
            {
                System.Console.WriteLine(ex.Message);
            }
        }

        public void checkReportDate(String text)
        {
            //find the RUNDATE and obtain the string value for the date.
            Regex pattern = new Regex("(RUNDATE[ 0-9]+/[0-9]+/[0-9]+)");
            MatchCollection mCollections = pattern.Matches(text);
            if (mCollections.Count > 0)
            {
                string datestr = mCollections[0].Value;
                //now we need to extract the date value from the datestr
                //we know that the string contains the text RUNDATE so we can get the substring value...
                datestr = datestr.Substring(7); //start at index 7 (length of Rundate)
                datestr = datestr.Trim();
                ReportDate(datestr);
            }
        }

        public void parseTransactionString(String inText, TransactionType inTransaction)
        {
            string str = inText;
            TransactionType aTransaction = inTransaction;
            Regex pattern = null;
            MatchCollection mCollection = null;


            //Need to parse out the individual values and stuff them into the member variables for ClientBranch.

            //get the transaction count...
            try
            {
                //get the Transaction Amount, Fee Amount, and Total Amount 
                pattern = new Regex("(\\s[0-9,.]+\\s)");
                mCollection = pattern.Matches(str);
                if (mCollection.Count > 0)
                {
                    aTransaction.TransactionCount = decimal.Parse(mCollection[0].Value);
                    aTransaction.TransactionAmount = decimal.Parse(mCollection[1].Value);
                    aTransaction.FeeAmount = decimal.Parse(mCollection[2].Value);
                    aTransaction.TotalAmount = decimal.Parse(mCollection[3].Value);
                }

                //get the CR|DR values...

                pattern = new Regex("(CR|DR)");
                mCollection = pattern.Matches(str);
                if (mCollection.Count > 0)
                {
                    //if we have 3 parsed values then we can insert them in order.
                    if (mCollection.Count == 3)
                    {
                        aTransaction.DBCR1 = mCollection[0].Value;
                        aTransaction.DBCR2 = mCollection[1].Value;
                        aTransaction.DBCR3 = mCollection[2].Value;
                    }

                    //if we have 2 then we need to determine where to insert the data
                    if (mCollection.Count == 2)
                    {
                        //if we have value here then we know that there are no Fees...
                        if (aTransaction.TransactionAmount > 0)
                        {
                            aTransaction.DBCR1 = mCollection[0].Value;
                            aTransaction.DBCR3 = mCollection[1].Value;
                        }
                        else //no value for TransactionAmount so we must have some sort of Fee...
                        {
                            aTransaction.DBCR2 = mCollection[0].Value;
                            aTransaction.DBCR3 = mCollection[1].Value;
                        }
                    }

                    //no need to check for one because you will at the very least have two because of the TOTAL Amouhnt Column!
                }
            }
            catch (Exception ex)
            {
                System.Console.WriteLine(ex.Message);
            }
            //Add the transaction in to the BranchInfo class...
            if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.MANUAL_ADJUSTMENT))
            {
                //need to check if transaction count is zero...
                if (aTransaction.TransactionCount == 0)
                {
                    return;
                }
            }

            //VS4731 4732 -- used for Card Activity Amount worksheet.
            if(aTransaction.Transaction.Equals(TransactionType.TransactionOption.LOAD_DISPUTES) ||
               aTransaction.Transaction.Equals(TransactionType.TransactionOption.PURCHASES_QUASI_CASH) ||
               aTransaction.Transaction.Equals(TransactionType.TransactionOption.PURCHASES_WITH_CASH_BACK) ||
               aTransaction.Transaction.Equals(TransactionType.TransactionOption.AFT_AA_PP) ||
               aTransaction.Transaction.Equals(TransactionType.TransactionOption.PURCHASE_RETURNS) ||
               aTransaction.Transaction.Equals(TransactionType.TransactionOption.MANUAL_CASH) ||
               aTransaction.Transaction.Equals(TransactionType.TransactionOption.ATM_CASH) ||
               aTransaction.Transaction.Equals(TransactionType.TransactionOption.BALANCE_INQUIRIES) ||
               aTransaction.Transaction.Equals(TransactionType.TransactionOption.REWARDS) ||
               aTransaction.Transaction.Equals(TransactionType.TransactionOption.TOTAL_CARD_ACTIVITY))
            {
                BranchInfo.CardActivityTransactionList.Add(aTransaction);
            }
            else
                BranchInfo.TransactionList.Add(aTransaction);
        }

        public void ResetData()
        {
            _branchInfo = new ClientBranch();
            _sarpage = 0;
            subPage = 1;
            _isglobalsummary = false;
            Data = "";
            ReportCancel = false;
            ReportFinish = false;
            isDateChecked = false;
        }
        #endregion
        #region new branch event

        public delegate void BranchChangedEventHandler(ClientBranch newBranch);
        public event BranchChangedEventHandler BranchChanged;
        #endregion

        #region end of report event

        public delegate void EndofReportEventHandler(String data);
        public event EndofReportEventHandler EndofReport;
        #endregion

        #region report date error notification
        public delegate void ReporDateEventHandler(String date);
        public event ReporDateEventHandler ReportDate;
        #endregion

    }
}
