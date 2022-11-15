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
/// Zusammenfassungsbeschreibung für TToken
/// </summary>
public class TToken
{
    public int ID;
    public string token;
    public bool valid;
    public DateTime timestamp;
    public string userID;
    public string toolID;
    public int age;

    /// <summary>
    /// Token aus der Datenbank lesen, auf Gültigkeit überprüfen und anschließend ggf. löschen. Das Token darf nicht älter als 10 Sekunden sein. Das erzeugte Token ist gültig wenn valid=true
    /// </summary>
    /// <param name="aToken">TokenID aus der URL</param>
    /// <param name="aToolID">ID des Tools für das das Token gültig sein soll</param>
    /// <param name="deleteToken">Flag zum Löschen des Token aus der Datenbank</param>
    public TToken(string aToken, string aToolID, bool deleteToken)
    {
        valid = false;
        bool foundToken = false;
        // aktuellen Zeitpunkt ermitteln
        DateTime actDateTime = DateTime.Now;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("token", "string", aToken);
        parameterList.addParameter("toolID", "string", aToolID);
        SqlDB dataReader;
        dataReader = new SqlDB("SELECT ID, token, timestamp, userID, toolID FROM token WHERE token=@token AND toolID=@toolID",parameterList, "");
        if (dataReader.read())
        {
            foundToken = true;
            ID = dataReader.getInt32(0);
            token = dataReader.getString(1);
            timestamp = dataReader.getDateTime(2);
            userID = dataReader.getString(3);
            toolID = dataReader.getString(4);
            valid = true;
        }
        dataReader.close();
        // Gültigkeit prüfen
        TimeSpan diff = actDateTime - timestamp;
        if (diff.TotalSeconds > 10)
        {
            valid = false;
        }
        // Token aus der Datenbank löschen
        if ((foundToken) & (deleteToken))
        {
            parameterList = new TParameterList();
            parameterList.addParameter("token", "int", this.ID.ToString());
            dataReader = new SqlDB("");
            dataReader.execSQLwithParameter("DELETE FROM token WHERE ID=@token",parameterList);
        }
    }
    /// <summary>
    /// Neues Token erstellen und in die Datenbank schreiben
    /// </summary>
    /// <param name="aUserID">eindeutige UserID des Benutzers für den das Token erzeugt wird</param>
    /// <param name="aToolID">eindeutige toolID der Applikation für die das Token erzeugt wird</param>
    public TToken(string aUserID, string aToolID)
    {
        token = System.Guid.NewGuid().ToString();
        timestamp = DateTime.Now;
        userID = aUserID;
        toolID = aToolID;

        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("token", "string", token);
        parameterList.addParameter("timestamp", "datetime", timestamp.ToString());
        parameterList.addParameter("userID", "string", aUserID);
        parameterList.addParameter("toolID", "string", aToolID);
        SqlDB dataReader = new SqlDB("");
        ID = dataReader.execSQLwithParameterAndReadID("INSERT INTO token (token, timestamp, userID, toolID) VALUES(@token,@timestamp,@userID,@toolID)",parameterList);
    }
}
