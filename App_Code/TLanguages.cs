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
/// Klasse zur Verwaltung der verfügbaren Sprachen der Tools
/// </summary>
public class TLanguages
{
    public class TEntry
    {
        public int ID;              // Primärschlüssel
        public string Language;    // Bezeichnung der Sprache
        public string Text;         // Anzeigetext der Sprache
        public string LanguageCode; // Iso-Code der Sprache
    }
    public ArrayList Language;

    /// <summary>
    /// Laden der Projektsprachen aus der DB
    /// </summary>
    /// <param name="aProject">ID des Projektes</param>
    public TLanguages(string aProject)
    {
        Language = new ArrayList();
        SqlDB dataReader;
        dataReader = new SqlDB("SELECT ID, language, displayName, languageCode FROM languages ORDER BY ID",aProject);
        while (dataReader.read())
        {
            TEntry TempEntry = new TEntry();
            TempEntry.ID = dataReader.getInt32(0);
            TempEntry.Language = dataReader.getString(1);
            TempEntry.Text = dataReader.getString(2);
            TempEntry.LanguageCode = dataReader.getString(3);
            Language.Add(TempEntry);
        }
        dataReader.close();
    }
    /// <summary>
    /// Abfrage ob eine Sprache für die Tools verfügbar ist
    /// </summary>
    /// <param name="aLanguage">zu überprüfende Sprache</param>
    /// <returns>true wenn Sprache verfügbar, sonst false</returns>
    public string getCodeByID(int aID)
    {
        string code = "";
        foreach (TEntry tempLanguage in Language)
            if (tempLanguage.ID == aID)
                code = tempLanguage.LanguageCode;
        return code; ;
    }
    public string getDisplaynameByID(int aID)
    {
        string displayname = "";
        foreach (TEntry tempLanguage in Language)
            if (tempLanguage.ID == aID)
                displayname = tempLanguage.Text;
        return displayname; ;
    }
    public bool exists(string aLanguage)
    {
        bool found = false;
        foreach (TEntry tempLanguage in Language)
            if (tempLanguage.Language == aLanguage)
                found = true;
        return found;
    }
    public static string getMasterLanguage(string aProject)
    {
        string result = "";
        SqlDB dataReader;
        dataReader = new SqlDB("SELECT language FROM languages ORDER BY ID", aProject);
        if (dataReader.read())
        {
            result = dataReader.getString(0);
        }
        dataReader.close();
        return result;
    }
}
