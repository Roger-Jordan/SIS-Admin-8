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
/// Zusammenfassungsbeschreibung für TAddValue1List
/// </summary>
public class TEmployeeCategoriesList
{
	public class TEntry
	{
        public int ID;          // Primärschlüssel
        public int Index;       // fortlaufende Nummer einer Kategorie im Array, beginnend mit 0
        public int Value;       // Schlüssel einer Kategorie, Sprachübergreifend gleich
        public string Text;     // Text einer Kategorie in einer Sprache
        public string Language; // Strache einer Kategorie
    }
    public int indexCounter;
    public int indexValue;
    public ArrayList ValueList;
    public int count;

	// erzeugt neues Objekt
    public TEmployeeCategoriesList(string aField, string aProject)
	{
        // Liste der Beteiligten erstellen
        indexCounter = -1;
        indexValue = -1;
        ValueList = new ArrayList();
        SqlDB dataReader;
        dataReader = new SqlDB("select ID, value, text, language from orgmanager_employeefieldscategories WHERE fieldID='" + aField + "' ORDER by value, language", aProject);
        while (dataReader.read())
        {
            if (indexValue != dataReader.getInt32(1))
            {
                indexCounter++;
                indexValue = dataReader.getInt32(1);
            }
            TEntry TempEntry = new TEntry();
            TempEntry.ID = dataReader.getInt32(0);
            TempEntry.Index = indexCounter;
            TempEntry.Value = dataReader.getInt32(1);
            TempEntry.Text = dataReader.getString(2);
            TempEntry.Language = dataReader.getString(3);
            ValueList.Add(TempEntry);
        }
        dataReader.close();
        dataReader = new SqlDB("select COUNT(DISTINCT value) FROM orgmanager_employeefieldscategories WHERE fieldID='" + aField + "'",aProject);
        count = 0;
        if (dataReader.read())
        {
            count = dataReader.getInt32(0);
        }
        dataReader.close();
    }
    public string getNameByValue(int aValue, string aLanguage)
    {
        string Result = "";
        int i;
        for (i = 0; i < ValueList.Count; i++)
        {
            if ((((TEntry)ValueList[i]).Value == aValue) & (((TEntry)ValueList[i]).Language == aLanguage))
            {
                Result = ((TEntry)ValueList[i]).Text;
            }
        }
        return Result;
    }

    public string getNameByIndex(int aIndex, string aLanguage)
    {
        string Result = "";
        int i;
        for (i = 0; i < ValueList.Count; i++)
        {
            if ((((TEntry)ValueList[i]).Index == aIndex) & (((TEntry)ValueList[i]).Language == aLanguage))
            {
                Result = ((TEntry)ValueList[i]).Text;
            }
        }
        return Result;
    }

    public int getValue(string aText)
    {
        int Result = -1;
        int i;
        for (i = 0; i < ValueList.Count; i++)
        {
            if (((TEntry)ValueList[i]).Text == aText)
            {
                Result = ((TEntry)ValueList[i]).Value;
            }
        }
        return Result;
    }

    public int getIndex(int aValue)
    {
        int Result = -1;
        int i;
        for (i = 0; i < ValueList.Count; i++)
        {
            if (((TEntry)ValueList[i]).Value == aValue)
            {
                Result = ((TEntry)ValueList[i]).Index;
            }
        }
        return Result;
    }
}
