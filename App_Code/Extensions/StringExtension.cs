using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SIS.Extensions
{
    /*
     * Erweiterung für String-Klasse, sodass sowas hier geht:
     * 
     *      string str = "a";
     *      str = str.AppendIfNotEmpty("b");
     *      // str == "a b"
     */
    public static class StringExtension
    {
        public static string AppendIfNotEmpty(this string str, string appendStr, string separator = " ")
        {
            if (string.IsNullOrEmpty(appendStr))
                return str;

            if (string.IsNullOrEmpty(str))
                return appendStr;

            str += separator + appendStr;
            return str;
        }
    }
}
