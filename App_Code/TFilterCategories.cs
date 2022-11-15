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
public class TFilterCategories
{
    public string FieldID;				// eindeutige Kennung des aktuellen Feldes
    public int Value;
    public int OrgID;
    public string RightType;

    /// <summary>
    /// Objekt erzeugen
    /// </summary>
    public TFilterCategories()
    {
    }
    /// <summary>
    /// Eigenschaften eines Benutzer aus der DB lesen
    /// </summary>
    /// <param name="aUserID">ID des Bentuzers</param>
    /// <param name="aProjectID">ID des Projektes</param>
    public TFilterCategories(string aTable, string aFieldID, int aValue, int aOrgID, string aProjectID)
    {
        SqlDB dataReader;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("fieldID", "string", aFieldID);
        parameterList.addParameter("value", "int", aValue.ToString());
        parameterList.addParameter("orgID", "int", aOrgID.ToString());
        dataReader = new SqlDB("select righttype FROM " + aTable + " WHERE fieldID=@fieldID AND value=@value AND orgID=@orgID", parameterList, aProjectID);
        if (dataReader.read())
        {
            FieldID = aFieldID;
            Value = aValue;
            OrgID = aOrgID;
            RightType = dataReader.getString(0);
        }
        dataReader.close();
    }
    /// <summary>
    /// Speichern eines neuen Benutzers in der DB
    /// </summary>
    /// <param name="aProjectID">ID des Projektes</param>
    /// <returns></returns>
    public bool save(string aTable, string aProjectID)
    {
        SqlDB dataReader;

        // fieldID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("fieldID", "string", this.FieldID);
        parameterList.addParameter("value", "int", this.Value.ToString());
        parameterList.addParameter("orgID", "int", this.OrgID.ToString());
        dataReader = new SqlDB("SELECT rightType FROM " + aTable + " WHERE fieldID=@fieldID AND value=@value AND orgID=@orgID", parameterList, aProjectID);
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
            parameterList.addParameter("orgID", "int", OrgID.ToString());
            parameterList.addParameter("rightType", "string", RightType);
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO " + aTable + " (fieldID, value, orgID, rightType) VALUES (@fieldID, @value, @orgID, @rightType)", parameterList);
            return true;
        }
        else
        {
            return false;
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
        parameterList.addParameter("value", "int", this.Value.ToString());
        parameterList.addParameter("orgID", "int", OrgID.ToString());
        parameterList.addParameter("rightType", "string", RightType);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("UPDATE " + aTable + " SET rightType=@rightType WHERE fieldID=@fieldID AND value=@value AND orgID=@orgID", parameterList);
    }
    /// <summary>
    /// Löschen eines Benutzers aus der DB
    /// </summary>
    /// <param name="aUserID">ID des Benutzers</param>
    /// <param name="aProjectID">ID des Projektes</param>
    public static void delete(string aTable, string aFieldID, int aValue, int aOrgID, string aProjectID)
    {
        SqlDB dataReader;
        // aus Datenbank löschen
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("fieldID", "string", aFieldID);
        parameterList.addParameter("value", "int", aValue.ToString());
        parameterList.addParameter("orgID", "int", aOrgID.ToString());
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM " + aTable + " WHERE fieldID=@fieldID AND value=@value AND orgID=@orgID", parameterList);
    }
}

