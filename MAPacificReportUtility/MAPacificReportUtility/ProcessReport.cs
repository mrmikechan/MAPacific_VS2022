using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace MAPacificReportUtility
{
    public abstract class ProcessReport
    {
        public ProcessReport()
        {
        }

        #region properties

        private int _columns;
        public int Columns
        {
            get { return _columns; }
            set
            {
                _columns = value;
            }
        }

        private int _rows;
        public int Rows
        {
            get { return _rows; }
            set
            {
                _rows = value;
            }
        }


        private Regex _regexPattern;
        public Regex RegexPattern
        {
            get { return _regexPattern; }
            set
            {
                if (value != null)
                {
                    _regexPattern = value;
                }
            }
        }

        private string _data;
        /// <summary>
        /// Data property to store any string data. Note that there is no trimming of the string in case
        /// formatting issues come into play.
        /// </summary>
        public string Data
        {
            get { return _data; }
            set
            {
                if (value != null)
                {
                    _data = value;
                }     
            }

        }

        private string _functionKey;
        public string FunctionKey
        {
            get { return _functionKey; }
            set
            {
                _functionKey = string.IsNullOrEmpty(value) ? "" : value.Trim();
            }
        }
        #endregion

        #region reg expression helper functions

        public bool isTargetTextPresent(string txt, Regex inPattern)
        {
            return inPattern.IsMatch(txt);
        }

        // Function to test for Positive Integers. 
        public bool IsNaturalNumber(String strNumber)
        {
            Regex objNotNaturalPattern = new Regex("[^0-9]");
            Regex objNaturalPattern = new Regex("0*[1-9][0-9]*");
            return !objNotNaturalPattern.IsMatch(strNumber) &&
            objNaturalPattern.IsMatch(strNumber);
        }
        // Function to test for Positive Integers with zero inclusive 
        public bool IsWholeNumber(String strNumber)
        {
            Regex objNotWholePattern = new Regex("[^0-9]");
            return !objNotWholePattern.IsMatch(strNumber);
        }
        // Function to Test for Integers both Positive & Negative 
        public bool IsInteger(String strNumber)
        {
            Regex objNotIntPattern = new Regex("[^0-9-]");
            Regex objIntPattern = new Regex("^-[0-9]+$|^[0-9]+$");
            return !objNotIntPattern.IsMatch(strNumber) && objIntPattern.IsMatch(strNumber);
        }
        // Function to Test for Positive Number both Integer & Real 
        public bool IsPositiveNumber(String strNumber)
        {
            Regex objNotPositivePattern = new Regex("[^0-9.]");
            Regex objPositivePattern = new Regex("^[.][0-9]+$|[0-9]*[.]*[0-9]+$");
            Regex objTwoDotPattern = new Regex("[0-9]*[.][0-9]*[.][0-9]*");
            return !objNotPositivePattern.IsMatch(strNumber) &&
            objPositivePattern.IsMatch(strNumber) &&
            !objTwoDotPattern.IsMatch(strNumber);
        }
        // Function to test whether the string is valid number or not
        public bool IsNumber(String strNumber)
        {
            Regex objNotNumberPattern = new Regex("[^0-9.-]");
            Regex objTwoDotPattern = new Regex("[0-9]*[.][0-9]*[.][0-9]*");
            Regex objTwoMinusPattern = new Regex("[0-9]*[-][0-9]*[-][0-9]*");
            String strValidRealPattern = "^([-]|[.]|[-.]|[0-9])[0-9]*[.]*[0-9]+$";
            String strValidIntegerPattern = "^([-]|[0-9])[0-9]*$";
            Regex objNumberPattern = new Regex("(" + strValidRealPattern + ")|(" + strValidIntegerPattern + ")");
            return !objNotNumberPattern.IsMatch(strNumber) &&
            !objTwoDotPattern.IsMatch(strNumber) &&
            !objTwoMinusPattern.IsMatch(strNumber) &&
            objNumberPattern.IsMatch(strNumber);
        }
        // Function To test for Alphabets. 
        public bool IsAlpha(String strToCheck)
        {
            Regex objAlphaPattern = new Regex("[^a-zA-Z]");
            return !objAlphaPattern.IsMatch(strToCheck);
        }

        //Function To test for Alphabets and space ("CALCOE FCU")
        public bool IsAlphaAndSpace(String strToCheck)
        {
            Regex objAlphaPattern = new Regex("[A-Za-z ]+");
            return objAlphaPattern.IsMatch(strToCheck);
        }
        // Function to Check for AlphaNumeric.
        public bool IsAlphaNumeric(String strToCheck)
        {
            Regex objAlphaNumericPattern = new Regex("[^a-zA-Z0-9]");
            return !objAlphaNumericPattern.IsMatch(strToCheck);
        }
        #endregion
    }
}
