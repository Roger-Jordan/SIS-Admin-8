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

/// <summary>
/// Klasse zur Verwaltung einer Parametrliste für parametrisierte Datenbankzugriffe
/// </summary>
public class TParameterList
{
    public class TEntry
    {
        public string name;       
        public string type;      
        public string value;     
    }

    public ArrayList parameter;

    /// <summary>
    /// Parameterliste erezeugen
    /// </summary>
    public TParameterList()
    {
        parameter = new ArrayList();
    }
    /// <summary>
    /// Parameter zur Parameterliste hinzufügen
    /// </summary>
    /// <param name="aName">Name des Paramters</param>
    /// <param name="aType">Typ des Parameters (int, float, datetime, string, text)</param>
    /// <param name="aValue">Inhalt/Wert des Parameters</param>
    public void addParameter(string aName, string aType, string aValue)
    {
        TEntry tempEntry = new TEntry();
        tempEntry.name = aName;
        tempEntry.type = aType;
        tempEntry.value = aValue;
        parameter.Add(tempEntry);
    }
    /// <summary>
    /// Ändern des Inhalts/Wertes eines Parameters
    /// </summary>
    /// <param name="aName">Name des zu ändernden Parameters</param>
    /// <param name="aValue">neuer Inhalt/Wert des Parameters</param>
    public void changeParameterValue(string aName, string aValue)
    {
        foreach(TEntry tempEntry in parameter)
        {
            if (tempEntry.name == aName)
                tempEntry.value = aValue;
        }
    }
}
