using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;

namespace MAPacificReportUtility
{
    /// <summary>
    /// PRCNOItem class object encapsulates the string value used for serializing the PRCNO from Wire Confirmation Report. Normally
    /// //one shouold be able to use a standard string however, the ultragrid does not like binding to a bindinglist<string>and thus this class was born.
    /// </summary>
    public class PRCNOItem : INotifyPropertyChanged
    {
        #region property
        private string prcno;
        public string PRCNO
        {
            get { return prcno; }
            set
            {
                prcno = string.IsNullOrEmpty(value) ? "" : value;
                OnPropertyChanged("PRCNO");
            }

        }
        #endregion

        public PRCNOItem(string inValue)
        {
            if (!string.IsNullOrEmpty(inValue))
            {
                PRCNO = inValue;
            }
        }

        /// <summary>
        /// PRCNOItem constructor is needed for the databinding by the Ultragrid. It make a call to a zero parameter constructor.
        /// If this is missing then the ultragrid complains when trying to add a new row data.
        /// </summary>
        public PRCNOItem()
        {

        }

        public event PropertyChangedEventHandler PropertyChanged;

        void OnPropertyChanged(string name)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(name));
            }
        }

    }
}
