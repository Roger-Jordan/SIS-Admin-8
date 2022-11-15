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
/// Klasse zur Verwaltung der Benutzers der Tools
/// </summary>
public class TResponseRight
{
    public int ID;              // eindeutige Kennung der Berechtigung
    public string UserID;		// ID des Benutzers		
    public int OrgID;           // ID der Org-Einheit, deren Rücklauf gesehen werden darf

    /// <summary>
    /// Objekt erzeugen
    /// </summary>
    public TResponseRight()
    {
    }
    /// <summary>
    /// Berechtigung aus DB lesen
    /// </summary>
    /// <param name="aID">ID der Berechtigung</param>
    /// <param name="aProjectID">ID des Projektes</param>
    public TResponseRight(int aID, string aProjectID)
    {
        ID = aID;
        SqlDB dataReader;
        dataReader = new SqlDB("select ID, userID, orgID from response_rights where ID='" + aID + "'", aProjectID);
        if (dataReader.read())
        {
            ID = dataReader.getInt32(0);
            UserID = dataReader.getString(1);
            OrgID = dataReader.getInt32(2);
        }
    }
    /// <summary>
    /// neue Berechtigung erzeugen
    /// </summary>
    /// <param name="aUserID">ID des Bentuzers</param>
    /// <param name="aOrgID">ID der Org-Einheit</param>
    /// <param name="aProjectID">ID des Projektes</param>
    public TResponseRight(string aUserID, int aOrgID, string aProjectID)
    {
        UserID = aUserID;
        OrgID = aOrgID;
    }
    /// <summary>
    /// Berechtigung in DB schreiben
    /// </summary>
    /// <param name="aProjectID"></param>
    public void save(string aProjectID)
    {
        SqlDB dataReader;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", UserID);
        parameterList.addParameter("orgID", "int", OrgID.ToString());
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("INSERT INTO response_rights (userID, orgID) VALUES (@userID, @orgID)", parameterList);
    }
    /// <summary>
    /// Berechtigung in DB aktunalisieren
    /// </summary>
    /// <param name="aProjectID">ID des Projektes</param>
    public void update(string aProjectID)
    {
        SqlDB dataReader;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", UserID);
        parameterList.addParameter("orgID", "int", OrgID.ToString());
        parameterList.addParameter("ID", "int", ID.ToString());
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("UPDATE response_rights SET userID=@userID, orgID=@orgID WHERE ID=@ID", parameterList);
    }
    /// <summary>
    /// Löschen eines Benutzers aus der DB
    /// </summary>
    /// <param name="aUserID">ID des Benutzers</param>
    /// <param name="aProjectID">ID des Projektes</param>
    public static void delete(string aID, string aProjectID)
    {
        SqlDB dataReader;
        // aus Datenbank löschen
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQL("DELETE FROM response_rights WHERE ID='" + aID + "'");
    }
}

