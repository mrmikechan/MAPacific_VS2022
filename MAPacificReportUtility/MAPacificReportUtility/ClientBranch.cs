using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Xml.Serialization;
///<remarks>
///ClientBranch class contains all the detailed information that are used to identify a branch. In addition
///to the id information, it will also contain data containers that will hold the parsed out data from the report.
///</remarks>

namespace MAPacificReportUtility
{
    public class ClientBranch : INotifyPropertyChanged, IEquatable<ClientBranch>
    {
       public ClientBranch()
       {
           _transactionlist = new BindingList<TransactionType>();
           _cardActivityList = new BindingList<TransactionType>();
       }

       #region datamembers

        private const string MAP_GPR_BIN = "99586001";
        private const string MAP_GIFT_BIN = "99586002";
        //VS4583 Enhance tool to work with a new entity named Beken which handles private bank reports.
        private const string BEKEN_GPR_BIN = "99818001";
        //There isn't one for Beken GIFT BIN yet...so we set arbitrary value of 0.
        private const string BEKEN_GIFT_BIN = "00000000";

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


       public struct ExcelSummaryType
       {
           public static string ALL
           {
               get { return "ALL"; }
           }

           public static string GPR_Gift_Summary_Prepaid_Total
           {
               get { return "GPR and Gift Summary and Prepaid Card Total"; }
           }

           public static string GPR_Gift_Details
           {
               get { return "GPR and Gift Details"; }
           }
       };

       //VS4583 Added Report to the xml config file.
       public struct ReportType
       {
           public static string BEKEN
           {
               get { return "BEKEN"; }
           }

           public static string MAP
           {
               get { return "MAP"; }
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

       //VS4142 Add bin to hold the BIN numbers for the two different GPR or GIFT report
       // so that we know which report this client belongs to!
       //VS4583 There are now 4 different BIN #'s
       private string _bin = "";

       public string BIN
       {
           get { return _bin; }
           set
           {
               if (!_bin.Equals(value))
               {
                   _bin = string.IsNullOrEmpty(value) ? "" : value.Trim();
                   OnPropertyChanged("BIN");
               }

               if ((_bin.Equals(MAP_GIFT_BIN)) || (_bin.Equals(MAP_GPR_BIN)))
               {
                   Report = ReportType.MAP;
               }

               if ((_bin.Equals(BEKEN_GIFT_BIN)) || (_bin.Equals(BEKEN_GPR_BIN)))
               {
                   Report = ReportType.BEKEN;
               }
           }
       }

       private string _clientid = "";
       /// <summary>
       /// string variable ClietID used to store the cliend id tag for this Branch.
       /// </summary>
       public string ClientID
       {
           get { return _clientid; }
           set
           {
               if (!_clientid.Equals(value))
               {
                   _clientid = string.IsNullOrEmpty(value) ? "" : value.Trim();
                   OnPropertyChanged("ClientID");
               }
           }
       }

        private string _name = "";
        /// <summary>
        ///string variable Name used to contain the name of the Branch 
        /// </summary>
        public string Name
        {
            get{ return _name;}
            set
            {
                if (!_name.Equals(value))
                {
                    _name = string.IsNullOrEmpty(value) ? "" : value.Trim();
                    OnPropertyChanged("Name");
                }
            }
           
        }



        private string _category = "";
        /// <summary>
        /// string variable Category used to contain the different possible categories that the client branch can be organized under:
        /// GPR Detail
        /// Gift Detail
        /// GPR Summary
        /// Gift Summary
        /// Ignore
        /// Prepaid Total Summary
        /// </summary>
        public string Category
        {
            get { return _category; }
            set
            {
                if (!_category.Equals(value))
                {
                    _category = string.IsNullOrEmpty(value) ? "" : value.Trim();
                    OnPropertyChanged("Category");
                }
            }
        }

        private string _excelsummary = "";
        /// <summary>
        /// string property ExcelSummary used to hold the various possible data outputs that customer branchs would like to receive. Note
        /// that the values for this property have not been flushed out yet. It will be done in the next phase of development.
        /// </summary>
        public string ExcelSummary
        {
            get { return _excelsummary; }
            set
            {
                if (!_excelsummary.Equals(value))
                {
                    _excelsummary = string.IsNullOrEmpty(value) ? "" : value.Trim();
                    OnPropertyChanged("ExcelSummary");
                }
            }
        }

        private BindingList<TransactionType> _transactionlist;
        [XmlIgnore]
        public BindingList<TransactionType> TransactionList
        {
            get { return _transactionlist; }
            set
            {
                if (value != null)
                {
                    _transactionlist = value;
                    OnPropertyChanged("TransactionList");
                }
            }
        }

        //VS4371 4372 add bindingist for cardactivity
        private BindingList<TransactionType> _cardActivityList;
        [XmlIgnore]
        public BindingList<TransactionType> CardActivityTransactionList
        {
            get { return _cardActivityList; }
            set
            {
                if (value != null)
                {
                    _cardActivityList = value;
                    OnPropertyChanged("CardActivityTransactionList");
                }
            }
        }

        private String _relationalParentSummary = "";
        public String RelationalParentSummary
        {
            get { return _relationalParentSummary; }
            set
            {
                if (value != null)
                {
                    _relationalParentSummary = string.IsNullOrEmpty(value) ? "" : value.Trim();
                    OnPropertyChanged("RelationalParentSummary");
                }
            }
        }

        private int _group;
        [XmlIgnore]
        public int Group
        {
            get { return _group; }
            set
            {
                _group = value;

                switch (_group)
                {
                    case ClientBranchBin.PrePaidCardsTotal:
                        {
                            Category = ClientBranch_Category.PREPAID_TOTAL_SUMMARY;
                        }
                        break;

                    case ClientBranchBin.GiftBranchDetail:
                        {
                            Category = ClientBranch_Category.GIFT_DETAIL;
                        }
                        break;

                    case ClientBranchBin.GiftCardSummary:
                        {
                            Category = ClientBranch_Category.GIFT_SUMMARY;
                        }
                        break;

                    case ClientBranchBin.GPRBranchDetail:
                        {
                            Category = ClientBranch_Category.GPR_DETAIL;
                        }
                        break;

                    case ClientBranchBin.GPRSummary:
                        {
                            Category = ClientBranch_Category.GPR_SUMMARY;
                        }
                        break;

                    case ClientBranchBin.IgnoreReport:
                        {
                            Category = ClientBranch_Category.IGNORE;
                        }
                        break;

                    case ClientBranchBin.UnrecognizedID:
                        {
                            Category = "";
                        }
                        break;
                            }
            }
        }

        //Report is used to associate the Client branch to the specific report. In our case, it will
        //either be MAP or Beken. This for the excel client reports so that we can display the pertinet customer
        //reports for Beken or Map instead of having both in there.
        private String _report = "";
        public String Report
        {
            get { return _report; }
            set
            {
                if (value != null)
                {
                    _report = string.IsNullOrEmpty(value) ? "" : value.Trim();
                    OnPropertyChanged("Report");
                }
            }
        }

        /// <summary>
        /// EmailAddress contains the email address for the CustomerID. Note that because the BranchInfo.xml keeps
        /// track of multiple ClientID's from same CreditUnion/Bank, you will see same emails for different entries here.
        /// Ideally we will only use the email address for ClientID that represents the Total Summary.
        /// </summary>
        private string _emailAddress = "";
        public string EmailAddress
        {
            get { return _emailAddress; }
            set
            {
                if (value != null)
                {
                    _emailAddress = string.IsNullOrEmpty(value) ? "" : value.Trim();
                    OnPropertyChanged("EmailAddress");
                }
            }
        }

        #endregion

        public bool Equals(ClientBranch other)
        {
            return this.ClientID.Equals(other.ClientID);
        }

        #region INotifyPropertyChanged Members

        public event PropertyChangedEventHandler PropertyChanged;

        void OnPropertyChanged(string name)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(name));
            }
        }
        #endregion

        #region utility function to clone clientBranch
        public ClientBranch Clone()
        {
            ClientBranch aCopy = (ClientBranch)this.MemberwiseClone();
            //now since memberwise clone performs a shallow copy we need to manually copy the BindingList...
            aCopy.TransactionList = new BindingList<TransactionType>();
            foreach (TransactionType mType in TransactionList)
            {
                aCopy.TransactionList.Add(mType.Clone());
            }
            return aCopy;
        }

        #endregion
    }
}
