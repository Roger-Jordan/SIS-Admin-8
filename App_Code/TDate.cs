using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Collections;
using System.Collections.Specialized;
using System.IO;
using System.Text;

/// <summary>
/// Klasse zur Verwaltung der Benutzers der Tools
/// </summary>
public class TDate
{
    public string UserID;				// eindeutige Kennung des aktuellen Benutzers
    public string DisplayName;	    	// Anzeigename des aktuellen Benutzers
    public string Password;		    	// Kennwort des aktuellen Benutzers
    public string Email;		    	// Email-Adresse der aktuellen Benutzers
    public string Title;
    public string Name;

    public static string getDate(string aString)
    {
        DateTime actDate;
        if (DateTime.TryParse(aString, System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None, out actDate))
        {
            return actDate.ToShortDateString();
        }
        else
        {
            int aInt;
            if (Int32.TryParse(aString, out aInt))
            {
                actDate = new DateTime(1899, 12, 31).AddDays(Convert.ToInt32(aInt) - 1);
                return actDate.ToShortDateString(); 
            }
            else
            {
                return "";
            }
        }
    }
}

