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
public class TCategorie
{
    public string FieldID;				
    public int Value;
    public string Text;
    public string Language;

    /// <summary>
    /// Objekt erzeugen
    /// </summary>
    public TCategorie()
    {
    }
    /// <summary>
    /// Eigenschaften eines Benutzer aus der DB lesen
    /// </summary>
    /// <param name="aUserID">ID des Bentuzers</param>
    /// <param name="aProjectID">ID des Projektes</param>
    public TCategorie(string aTable, string aFieldID, int aValue, string aLanguage, string aProjectID)
    {
        SqlDB dataReader;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("fieldID", "string", aFieldID);
        parameterList.addParameter("value", "int", aValue.ToString());
        parameterList.addParameter("language", "string", aLanguage);
        dataReader = new SqlDB("select text FROM " + aTable + " WHERE fieldID=@fieldID AND value=@value AND language=@language", parameterList, aProjectID);
        if (dataReader.read())
        {
            FieldID = aFieldID;
            Value = aValue;
            Language = aLanguage;
            Text = dataReader.getString(0);
        }
        dataReader.close();
    }
    /// <summary>
    /// Speichern einer Kategorie, in allen Sprachen soweit diese noch nciht bereits vorhanden sind
    /// </summary>
    /// <param name="aProjectID">ID des Projektes</param>
    /// <returns></returns>
    public void save(string aTable, string aProjectID)
    {
        SqlDB dataReader;

        // categorie auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("fieldID", "string", this.FieldID);
        parameterList.addParameter("value", "int", Value.ToString());
        parameterList.addParameter("language", "string", Language);
        dataReader = new SqlDB("SELECT userID FROM " + aTable + " WHERE fieldID=@fieldID AND value=@value AND language=@language", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            // Schleife über alle Sprachen und in allen noch nicht vorhandenen Sprachen einfügen
            TLanguages LanguageList = new TLanguages(aProjectID);
            for (int Index = 0; Index < LanguageList.Language.Count; Index++)
            {
                string actLanguage = ((TLanguages.TEntry)LanguageList.Language[Index]).Language;

                if (actLanguage == Language)
                {
                    // aktuelle Sprache
                    parameterList = new TParameterList();
                    parameterList.addParameter("fieldID", "string", FieldID);
                    parameterList.addParameter("value", "int", Value.ToString());
                    parameterList.addParameter("language", "string", Language);
                    parameterList.addParameter("text", "string", Text);
                    dataReader = new SqlDB(aProjectID);
                    dataReader.execSQLwithParameter("INSERT INTO " + aTable + "categories (fieldID, value, language, text) VALUES (@fieldID, @value, @language, @text)", parameterList);
                }
                else
                {
                    // andere Sprache auf vorhandensein prüfen und ggf. bei Nichtvorhandensein einfügen
                    parameterList = new TParameterList();
                    parameterList.addParameter("fieldID", "string", this.FieldID);
                    parameterList.addParameter("value", "int", Value.ToString());
                    parameterList.addParameter("language", "string", actLanguage);
                    dataReader = new SqlDB("SELECT userID FROM " + aTable + " WHERE fieldID=@fieldID AND value=@value AND language=@language", parameterList, aProjectID);
                    if (dataReader.read())
                    {
                        exists = true;
                    }
                    dataReader.close();
                    if (!exists)
                    {
                        parameterList = new TParameterList();
                        parameterList.addParameter("fieldID", "string", FieldID);
                        parameterList.addParameter("value", "int", Value.ToString());
                        parameterList.addParameter("language", "string", actLanguage);
                        parameterList.addParameter("text", "string", "-undefined-");
                        dataReader = new SqlDB(aProjectID);
                        dataReader.execSQLwithParameter("INSERT INTO " + aTable + "categories (fieldID, value, language, text) VALUES (@fieldID, @value, @language, @text)", parameterList);
                    }
                }
            }
        }
    }
    /// <summary>
    /// Aktualisieren der Eigenschaften eines Benutzers in der DB
    /// </summary>
    /// <param name="aProjectID">ID des Projektes</param>
    public void update(string aTable, string aProjectID)
    {
        SqlDB dataReader;

        TParameterList parameterList = new TParameterList();
        parameterList = new TParameterList();
        parameterList.addParameter("fieldID", "string", FieldID);
        parameterList.addParameter("value", "int", Value.ToString());
        parameterList.addParameter("language", "string", Language);
        parameterList.addParameter("text", "string", Text);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("UPDATE " + aTable + "categories SET text=@text WHERE fieldID=@fieldID AND value=@value AND language=@language", parameterList);
    }
    /// <summary>
    /// Löschen eines Benutzers aus der DB
    /// </summary>
    /// <param name="aUserID">ID des Benutzers</param>
    /// <param name="aProjectID">ID des Projektes</param>
    public static void delete(string aTable, string aFieldID, int aValue, string aProjectID)
    {
        SqlDB dataReader;
        // aus Datenbank löschen
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("fieldID", "string", aFieldID);
        parameterList.addParameter("value", "int", aValue.ToString());
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM " + aTable + " WHERE fieldID=@fieldID AND value=@value", parameterList);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM " + aTable + "filter WHERE fieldID=@fieldID AND value=@value", parameterList);
    }
}

