using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Xml.Serialization;
using System.Linq;

namespace MAPacificReportUtility
{
    //VS4229 created a class object to encapsulate dataobjects that we want to serialize to the file branchinfo.xml
    /// <summary>
    /// DataContainer class contains the objects that we want to serialize to a file. Normally the XMLSerializer can only serialize
    /// one class type at a time. As a workaround, we reference multiple objects that we want into a single xml file.
    /// </summary>
    public class DataContainer : INotifyPropertyChanged
    {
        #region properties
        private BindingList<ClientBranch> clientbranchlist;
        public  BindingList<ClientBranch> ClientBranchList
        {
            get { return clientbranchlist; }
            set 
            {
                if(value != null)
                {
                    clientbranchlist = value;
                    OnPropertyChanged("ClientBranchList");
                }

            }
        }

        /// <summary>
        /// XMLSerializer can not serialize a dictionary, so as a workaround we can
        /// serialize a list of KeyValuePair.
        /// </summary>
        private BindingList<PRCNOItem> wireconftablelist;
        public BindingList<PRCNOItem> WireConfTableList
        {
            get { return wireconftablelist; }
            set
            {
                if (value != null)
                {
                    wireconftablelist = value;
                    OnPropertyChanged("WireConfTableList");
                }
            }
        }
        #endregion

        public DataContainer()
        {
            ClientBranchList = new BindingList<ClientBranch>();
            WireConfTableList = new BindingList<PRCNOItem>();
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

        #region utility function
        public Dictionary<string, bool> GetDictionaryFromListKeyPair()
        {
            Dictionary<string, bool> myDictionary = null;
            //reassemble the list into a dictionary for use.
            if (WireConfTableList != null)
            {
                //myDictionary = WireConfTableList.ToDictionary(v => v, v => true);
                myDictionary = new Dictionary<string, bool>();
                foreach (PRCNOItem item in WireConfTableList)
                {
                    myDictionary.Add(item.PRCNO, true);
                }
            }
            return myDictionary;
        }
        #endregion
    }
}
