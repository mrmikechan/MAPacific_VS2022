using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using Infragistics.Win.UltraWinGrid;
using Infragistics.Win;

namespace MAPacificReportUtility.excel
{
    public class ExportExcelCardActivityCount
    {
        private List<TransactionCardCountItem> mtrnctlist;
        public List<TransactionCardCountItem> TransactionCardCountList
        {
            get { return mtrnctlist; }
            set
            {
                if (value != null)
                {
                    mtrnctlist = value;
                }
            }
        }

        private TransactionCardCountItem mTotalItem;
        public TransactionCardCountItem TransactionTotalItem
        {
            get { return mTotalItem; }
        }

        private UltraGrid ultragridCardCount;
        public UltraGrid UltgragridCardCount
        {
            get { return ultragridCardCount; }
        }

        //decimal objects used to hold running total column data...
        private decimal mTotal_Load_Disputes;
        private decimal mTotal_Purchases_Quasi_Cash;
        private decimal mTotal_Purchases_With_Cash_Back;
        private decimal mTotal_AFT_AA_PP;
        private decimal mTotal_Purchase_Returns;
        private decimal mTotal_Manual_Cash;
        private decimal mTotal_Atm_Cash;
        private decimal mTotal_Balance_Inquiries;
        private decimal mTotal_Rewards;
        private decimal mTotal_Total_Card_Activity;

        /// <summary>
        /// ExportExcelTransationAmount constructor.
        /// </summary>
        /// <param name="prepaidTotalSummary">BindingList of type ClientBranch that contains data for the Pre Paid Cards Total Summary</param>
        /// <param name="isBeken">bool isBeken used to determine if report is Beken or Map report type</param>
        public ExportExcelCardActivityCount(System.ComponentModel.BindingList<ClientBranch> prepaidTotalSummary, bool isBeken)
        {
            TransactionCardCountList = new List<TransactionCardCountItem>();
            mTotal_Load_Disputes = 0;
            mTotal_Purchases_Quasi_Cash = 0;
            mTotal_Purchases_With_Cash_Back = 0;
            mTotal_AFT_AA_PP = 0;
            mTotal_Purchase_Returns = 0;
            mTotal_Manual_Cash = 0;
            mTotal_Atm_Cash = 0;
            mTotal_Balance_Inquiries = 0;
            mTotal_Rewards = 0;
            mTotal_Total_Card_Activity = 0;

            ultragridCardCount = new UltraGrid();
            ultragridCardCount.InitializeRow += new InitializeRowEventHandler(ultragridCardCount_InitializeRow);
            ultragridCardCount.InitializeLayout += new InitializeLayoutEventHandler(ultragridCardCount_InitializeLayout);

            //Set the font to Arial.
            FontData fd1 = ultragridCardCount.DisplayLayout.Appearance.FontData;
            fd1.Name = "Arial";

            ultragridCardCount.BindingContext = new System.Windows.Forms.BindingContext();
            ProcessList(prepaidTotalSummary, isBeken);
            SetColumnHeaderLayout();
        }

        /// <summary>
        /// InitializeLayout event setup up the text alignment in the cells for each column when exported out to Excel.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ultragridCardCount_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            e.Layout.Bands[0].Columns["CLIENT_ID"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["CREDIT_UNION"].PerformAutoResize(Infragistics.Win.UltraWinGrid.PerformAutoSizeType.AllRowsInBand);
            e.Layout.Bands[0].Columns["LOAD_DISPUTES"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["PURCHASES_QUASI_CASH"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["PURCHASES_WITH_CASH_BACK"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["AFT_AA_PP"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["PURCHASE_RETURNS"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["MANUAL_CASH"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["ATM_CASH"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["BALANCE_INQUIRIES"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["REWARDS"].CellAppearance.TextHAlign = HAlign.Center;
            e.Layout.Bands[0].Columns["TOTAL_CARD_ACTIVITY"].CellAppearance.TextHAlign = HAlign.Center;

        }

        public void SetColumnHeaderLayout()
        {
            //Initialize the header columns in the Ultragrid to a specific font, bold, height, and color.  These settings
            //from the column header will then be used when exported into a Excel worksheet.
            foreach (UltraGridColumn col in ultragridCardCount.DisplayLayout.Bands[0].Columns)
            {
                col.Header.Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
                col.Header.Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                col.Header.Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                col.Header.Appearance.BackGradientStyle = GradientStyle.Vertical;
                col.Header.Appearance.FontData.Bold = DefaultableBoolean.True;
                col.Header.Appearance.FontData.SizeInPoints = 10;
                col.Header.Appearance.FontData.Name = "Arial";
                col.Header.Appearance.ForeColor = System.Drawing.Color.White;
                col.Header.Appearance.TextHAlign = HAlign.Center;
            }
        }


        void ultragridCardCount_InitializeRow(object sender, InitializeRowEventArgs e)
        {
            //Highlight the Total row in the grid so that when exported out to Excel that row is also highlighted.
            if (e.Row.Cells["CLIENT_ID"].Value.ToString().Equals("Total", StringComparison.CurrentCultureIgnoreCase))
            {
                e.Row.Cells["CLIENT_ID"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["CLIENT_ID"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["CLIENT_ID"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["CLIENT_ID"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["CLIENT_ID"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["CLIENT_ID"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["CLIENT_ID"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["CREDIT_UNION"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["CREDIT_UNION"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["CREDIT_UNION"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["CREDIT_UNION"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["CREDIT_UNION"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["CREDIT_UNION"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["CREDIT_UNION"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["LOAD_DISPUTES"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["LOAD_DISPUTES"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["LOAD_DISPUTES"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["LOAD_DISPUTES"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["LOAD_DISPUTES"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["LOAD_DISPUTES"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["LOAD_DISPUTES"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["PURCHASES_QUASI_CASH"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["PURCHASES_QUASI_CASH"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["PURCHASES_QUASI_CASH"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["PURCHASES_QUASI_CASH"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["PURCHASES_QUASI_CASH"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["PURCHASES_QUASI_CASH"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["PURCHASES_QUASI_CASH"].Appearance.ForeColor = System.Drawing.Color.White;


                e.Row.Cells["PURCHASES_WITH_CASH_BACK"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["PURCHASES_WITH_CASH_BACK"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["PURCHASES_WITH_CASH_BACK"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["PURCHASES_WITH_CASH_BACK"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["PURCHASES_WITH_CASH_BACK"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["PURCHASES_WITH_CASH_BACK"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["PURCHASES_WITH_CASH_BACK"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["AFT_AA_PP"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["AFT_AA_PP"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["AFT_AA_PP"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["AFT_AA_PP"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["AFT_AA_PP"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["AFT_AA_PP"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["AFT_AA_PP"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["PURCHASE_RETURNS"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["PURCHASE_RETURNS"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["PURCHASE_RETURNS"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["PURCHASE_RETURNS"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["PURCHASE_RETURNS"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["PURCHASE_RETURNS"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["PURCHASE_RETURNS"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["MANUAL_CASH"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["MANUAL_CASH"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["MANUAL_CASH"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["MANUAL_CASH"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["MANUAL_CASH"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["MANUAL_CASH"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["MANUAL_CASH"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["ATM_CASH"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["ATM_CASH"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["ATM_CASH"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["ATM_CASH"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["ATM_CASH"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["ATM_CASH"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["ATM_CASH"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["BALANCE_INQUIRIES"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["BALANCE_INQUIRIES"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["BALANCE_INQUIRIES"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["BALANCE_INQUIRIES"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["BALANCE_INQUIRIES"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["BALANCE_INQUIRIES"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["BALANCE_INQUIRIES"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["REWARDS"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["REWARDS"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["REWARDS"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["REWARDS"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["REWARDS"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["REWARDS"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["REWARDS"].Appearance.ForeColor = System.Drawing.Color.White;

                e.Row.Cells["TOTAL_CARD_ACTIVITY"].Appearance.BackColor = System.Drawing.Color.FromArgb(89, 135, 214);
                e.Row.Cells["TOTAL_CARD_ACTIVITY"].Appearance.BackColor2 = System.Drawing.Color.FromArgb(7, 59, 150);
                e.Row.Cells["TOTAL_CARD_ACTIVITY"].Appearance.BackGradientStyle = GradientStyle.Vertical;
                e.Row.Cells["TOTAL_CARD_ACTIVITY"].Appearance.FontData.Bold = DefaultableBoolean.True;
                e.Row.Cells["TOTAL_CARD_ACTIVITY"].Appearance.FontData.SizeInPoints = 10;
                e.Row.Cells["TOTAL_CARD_ACTIVITY"].Appearance.FontData.Name = "Arial";
                e.Row.Cells["TOTAL_CARD_ACTIVITY"].Appearance.ForeColor = System.Drawing.Color.White;
            }
        }

        /// <summary>
        /// ProcessList takes the BindingList of type ClientBranch and processes the data in there and massage them into a new List that
        /// is used to generate a new excel worksheet for Transaction Amount
        /// </summary>
        /// <param name="inList">List of ClientBranch types that contains the data needed to be processed</param>
        /// <param name="isBeken">bool isBeken used to determine if report is Beken or Map report type</param>
        private void ProcessList(System.ComponentModel.BindingList<ClientBranch> inList, bool isBeken)
        {

            TransactionCardCountItem newItem;
            foreach (ClientBranch item in inList)
            {
                if (isBeken)
                {
                    //Beken report type. We only want to see client ID related to Beken reports.
                    if (!item.Report.Equals(ClientBranch.ReportType.BEKEN))
                        continue;
                }

                if (!isBeken)
                {
                    if (item.Report.Equals(ClientBranch.ReportType.BEKEN))
                        continue;
                }

                newItem = new TransactionCardCountItem();
                newItem.CLIENT_ID = item.ClientID;
                newItem.CREDIT_UNION = item.Name;

                //look into the list of transactions in the ClientBranch
                //add the data into the new item and also calculate a running total value for each column...
                if (item.CardActivityTransactionList.Count > 0)
                {
                    foreach (TransactionType aTransaction in item.CardActivityTransactionList)
                    {
                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.LOAD_DISPUTES))
                        {
                            newItem.LOAD_DISPUTES = aTransaction.TransactionCount;
                            mTotal_Load_Disputes += aTransaction.TransactionCount;
                        }

                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.PURCHASES_QUASI_CASH))
                        {
                            newItem.PURCHASES_QUASI_CASH = aTransaction.TransactionCount;
                            mTotal_Purchases_Quasi_Cash += aTransaction.TransactionCount;
                        }

                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.PURCHASES_WITH_CASH_BACK))
                        {
                            newItem.PURCHASES_WITH_CASH_BACK = aTransaction.TransactionCount;
                            mTotal_Purchases_With_Cash_Back += aTransaction.TransactionCount;
                        }

                        //VS4637 Add Loads Bypass
                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.AFT_AA_PP))
                        {
                            newItem.AFT_AA_PP = aTransaction.TransactionCount;
                            mTotal_AFT_AA_PP += aTransaction.TransactionCount;
                        }

                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.PURCHASE_RETURNS))
                        {
                            newItem.PURCHASE_RETURNS = aTransaction.TransactionCount;
                            mTotal_Purchase_Returns += aTransaction.TransactionCount;
                        }

                        //VS4637 Add Loads Merchant Acq
                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.MANUAL_CASH))
                        {
                            newItem.MANUAL_CASH = aTransaction.TransactionCount;
                            mTotal_Manual_Cash += aTransaction.TransactionCount;
                        }

                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.ATM_CASH))
                        {
                            newItem.ATM_CASH = aTransaction.TransactionCount;
                            mTotal_Atm_Cash += aTransaction.TransactionCount;
                        }

                        //VS4637 Add Unloads Merchant Acq
                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.BALANCE_INQUIRIES))
                        {
                            newItem.BALANCE_INQUIRIES = aTransaction.TransactionCount;
                            mTotal_Balance_Inquiries += aTransaction.TransactionCount;
                        }

                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.REWARDS))
                        {
                            newItem.REWARDS = aTransaction.TransactionCount;
                            mTotal_Rewards += aTransaction.TransactionCount;
                        }

                        if (aTransaction.Transaction.Equals(TransactionType.TransactionOption.TOTAL_CARD_ACTIVITY))
                        {
                            newItem.TOTAL_CARD_ACTIVITY = aTransaction.TransactionCount;
                            mTotal_Total_Card_Activity += aTransaction.TransactionCount;
                        }
                    }
                }

                TransactionCardCountList.Add(newItem);
            }
            //finished processing the BindingList... now lets create a new row to add the total columns into the new list...
            mTotalItem = new TransactionCardCountItem();
            mTotalItem.CLIENT_ID = "Total";
            mTotalItem.CREDIT_UNION = "";
            mTotalItem.LOAD_DISPUTES = mTotal_Load_Disputes;
            mTotalItem.PURCHASES_QUASI_CASH = mTotal_Purchases_Quasi_Cash;
            mTotalItem.PURCHASES_WITH_CASH_BACK = mTotal_Purchases_With_Cash_Back;
            mTotalItem.AFT_AA_PP = mTotal_AFT_AA_PP;
            mTotalItem.PURCHASE_RETURNS = mTotal_Purchase_Returns;
            mTotalItem.MANUAL_CASH = mTotal_Manual_Cash;
            mTotalItem.ATM_CASH = mTotal_Atm_Cash;
            mTotalItem.BALANCE_INQUIRIES = mTotal_Balance_Inquiries;
            mTotalItem.REWARDS = mTotal_Rewards;
            //VS4637 Add the values from Load Bypass, Loads Merchant Acq, and Unloads Merchant Acq
            mTotalItem.TOTAL_CARD_ACTIVITY = mTotal_Total_Card_Activity;
            TransactionCardCountList.Add(mTotalItem);
            ultragridCardCount.DataSource = TransactionCardCountList;
        }
    }

    //Note that the order of properties listed is important.
    //They are displayed in this specific order when exported into Excel.
    public class TransactionCardCountItem : INotifyPropertyChanged
    {

        private string mclientID = "";
        public string CLIENT_ID
        {
            get { return mclientID; }
            set
            {
                if (!mclientID.Equals(value))
                {
                    mclientID = string.IsNullOrEmpty(value) ? "" : value.Trim();
                    OnPropertyChanged("CLIENT_ID");
                }
            }
        }

        private string mcreditUnion = "";
        public string CREDIT_UNION
        {
            get { return mcreditUnion; }
            set
            {
                if (!mcreditUnion.Equals(value))
                {
                    mcreditUnion = string.IsNullOrEmpty(value) ? "" : value.Trim();
                    OnPropertyChanged("CREDIT_UNION");
                }
            }
        }


        private decimal mloadsdisputes = 0;
        public decimal LOAD_DISPUTES
        {
            get { return mloadsdisputes; }
            set
            {
                if (mloadsdisputes != value)
                {
                    mloadsdisputes = value;
                    OnPropertyChanged("LOAD_DISPUTES");
                }
            }
        }

        private decimal mpurchasesquasicash = 0;
        public decimal PURCHASES_QUASI_CASH
        {
            get { return mpurchasesquasicash; }
            set
            {
                if (mpurchasesquasicash != value)
                {
                    mpurchasesquasicash = value;
                    OnPropertyChanged("PURCHASES_QUASI_CASH");
                }
            }
        }

        private decimal mpurchaseswithcashback = 0;
        public decimal PURCHASES_WITH_CASH_BACK
        {
            get { return mpurchaseswithcashback; }
            set
            {
                if (mpurchaseswithcashback != value)
                {
                    mpurchaseswithcashback = value;
                    OnPropertyChanged("PURCHASES_WITH_CASH_BACK");
                }
            }
        }

        private decimal maftaapp = 0;
        public decimal AFT_AA_PP
        {
            get { return maftaapp; }
            set
            {
                if (maftaapp != value)
                {
                    maftaapp = value;
                    OnPropertyChanged("AFT_AA_PP");
                }
            }
        }

        //VS4637
        private decimal mpurchasereturns = 0;
        public decimal PURCHASE_RETURNS
        {
            get { return mpurchasereturns; }
            set
            {
                if (mpurchasereturns != value)
                {
                    mpurchasereturns = value;
                    OnPropertyChanged("PURCHASE_RETURNS");
                }
            }
        }

        private decimal mmanualcash = 0;
        public decimal MANUAL_CASH
        {
            get { return mmanualcash; }
            set
            {
                if (mmanualcash != value)
                {
                    mmanualcash = value;
                    OnPropertyChanged("MANUAL_CASH");
                }
            }
        }

        //VS4637
        private decimal matmcash = 0;
        public decimal ATM_CASH
        {
            get { return matmcash; }
            set
            {
                if (matmcash != value)
                {
                    matmcash = value;
                    OnPropertyChanged("ATM_CASH");
                }
            }
        }
        private decimal mbalanceinquiries = 0;
        public decimal BALANCE_INQUIRIES
        {
            get { return mbalanceinquiries; }
            set
            {
                if (mbalanceinquiries != value)
                {
                    mbalanceinquiries = value;
                    OnPropertyChanged("BALANCE_INQUIRIES");
                }
            }
        }

        //VS4637
        private decimal mrewards = 0;
        public decimal REWARDS
        {
            get { return mrewards; }
            set
            {
                if (mrewards != value)
                {
                    matmcash = value;
                    OnPropertyChanged("REWARDS");
                }
            }
        }

        private decimal mtotalcardactivity = 0;
        public decimal TOTAL_CARD_ACTIVITY
        {
            get { return mtotalcardactivity; }
            set
            {
                if (mtotalcardactivity != value)
                {
                    mtotalcardactivity = value;
                    OnPropertyChanged("TOTAL_CARD_ACTIVITY");
                }
            }
        }


        #region INotifyPropertyChanged Members

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;

        void OnPropertyChanged(string name)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new System.ComponentModel.PropertyChangedEventArgs(name));
            }
        }
        #endregion

    }
}
