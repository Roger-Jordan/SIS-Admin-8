using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

/// <summary>
/// Klasse zum eintragen von Login und erfolglosen Login in die DB
/// </summary>
public class Login
{
    /// <summary>
    /// Protokollieren eines erfolgreichen Login
    /// </summary>
    /// <param name="aProjectID">ID des Projektes</param>
    /// <param name="aUserID">ID des Benutzers</param>
    /// <param name="aPassword">Passwort des Benutzers</param>
    public static void LoginMessage(string aProjectID, string aUserID, string aPassword)
    {
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("projectID", "string", aProjectID);
        parameterList.addParameter("userID", "string", aUserID);
        parameterList.addParameter("password", "string", "");
        parameterList.addParameter("datetime", "datetime", DateTime.Now.ToString());
        SqlDB dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("INSERT INTO login (projectID, userID, password, datetime) values (@projectID,@userID,@password,@datetime)",parameterList);
    }
    /// <summary>
    /// Protokollieren eines erfolglosen Login
    /// </summary>
    /// <param name="aProjectID">ID des Projektes</param>
    /// <param name="aUserID">verwendete ID des Benutzers</param>
    /// <param name="aPassword">verwendetes Passwort des Benutzers</param>
    public static void LoginErrorMessage(string aProjectID, string aUserID, string aPassword)
    {
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("projectID", "string", aProjectID);
        parameterList.addParameter("userID", "string", aUserID);
        parameterList.addParameter("password", "string", "");
        parameterList.addParameter("datetime", "datetime", DateTime.Now.ToString());
        SqlDB dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("INSERT INTO loginerror (projectID, userID, password, datetime) values (@projectID,@userID,@password,@datetime)", parameterList);
    }
}
