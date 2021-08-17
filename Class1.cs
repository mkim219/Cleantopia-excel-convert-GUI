using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Conversion_UI
{
    public class FilteredData
    {
        public string one_digit { get; set; }
        public string three_digit { get; set; }
        public string product { get; set; }
        public string date { get; set; }
        public string name { get; set; }

        public FilteredData(string _one_digit, string _three_digit, string _product, string _date, string _name)
        {
            one_digit = _one_digit;
            three_digit = _three_digit;
            product = _product;
            date = _date;
            name = _name;
        }
    }

    public class ExtractedData
    {   
        public string combinedDigits { get; set; }
        public string product { get; set; }
        public string date { get; set; }
        public string name { get; set; }

        public ExtractedData(string _combineDigits, string _product, string _date, string _name)
        {
            combinedDigits = _combineDigits;
            product = _product;
            date = _date;
            name = _name;
        }
    }
}
