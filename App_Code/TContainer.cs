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
/// Zusammenfassungsbeschreibung für TStructure
/// </summary>
public class TContainer
{
    public int containerID;			            // eindeutige ID des Containers
    public string displayName;	    	// Anzeigetext des Containers
    public int maxSizeGB;               // maximale Größe der Containers in GigaByte
    public int maxFileSizeMB;           // maximale Größe einer Datei im Container in MegaByte
    public bool passwordAuto;           // Kennzeichen, ob jede Datei im Container automatisch mit einem Kennwort geschützt werden soll
    public bool passwordMandatory;      // Kennzeichen, ob jede Datei im Container durch den Uploader mit einem Kennwort geschützt werden muss
    public bool passwordOptional;       // Kennzeichen, on eine Datei im Container durch den Uploader mit einem Kennwort geschützt werden kann
    public bool historyAuto;            // Kennzeichen, ob eine automatische Versionsverwaltung durchgeführt werden soll
    public bool historyOptional;        // Kennzeichen, ob eine Versionsverwaltung einer Datei durchgeführt werden kann
    public string path;                 // Pfad des Containers im Dateisystem

    /// <summary>
    /// Einlesen der Daten eines Containers
    /// </summary>
    /// <param name="aOrgID"></param>
    /// <param name="aParent"></param>
    /// <param name="aProject"></param>
    /// <param name="aAccessList"></param>
    /// <param name="aReadlevel"></param>
    public TContainer(int aContainerID, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        SqlDB dataReader;
        dataReader = new SqlDB("SELECT displayName, maxSizeGB, maxFileSizeMB, passwordAuto, passwordMandatory, passwordOptional, historyAuto, historyOptional FROM teamspace_container WHERE containerID='" + aContainerID.ToString() + "'", aProjectID);
        if (dataReader.read())
        {
            containerID = aContainerID;
            displayName = dataReader.getString(0);
            maxSizeGB = dataReader.getInt32(1);
            maxFileSizeMB = dataReader.getInt32(2);
            passwordAuto = dataReader.getBool(3);
            passwordMandatory = dataReader.getBool(4);
            passwordOptional = dataReader.getBool(5);
            historyAuto = dataReader.getBool(6);
            historyOptional = dataReader.getBool(7);
            path = SessionObj.ProjectDataPath + SessionObj.Project.ProjectID + "\\TeamSpace\\Files\\" + displayName;
        }
        dataReader.close();
    }
    public static string getContainerName(int aContainerID, string aProjectID)
    {
        string result = "";
        
        SqlDB dataReader;
        dataReader = new SqlDB("SELECT displayname FROM teamspace_container WHERE containerID='" + aContainerID + "'", aProjectID);
        if (dataReader.read())
        {
            result = dataReader.getString(0);
        }
        dataReader.close();

        return result;
    }
    public void update(string aDisplayName, int aMaxSizeGB, int aMaxFileSizeMB, bool aPasswortAuto, bool aPasswortMandatory, bool aPasswordOptional, bool aHistoryAuto, bool aHistoryOptional)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        displayName = aDisplayName;
        maxSizeGB = aMaxSizeGB;
        maxFileSizeMB = aMaxFileSizeMB;
        passwordAuto = aPasswortAuto;
        passwordMandatory = aPasswortMandatory;
        passwordOptional = aPasswordOptional;
        historyAuto = aHistoryAuto;
        historyOptional = aHistoryOptional;

        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("containerID", "int", containerID.ToString());
        parameterList.addParameter("displayName", "string", displayName);
        parameterList.addParameter("maxSizeGB", "int", maxSizeGB.ToString());
        parameterList.addParameter("maxFileSizeMB", "int", maxFileSizeMB.ToString());
        string tempPasswordAuto = "0";
        if (passwordAuto)
            tempPasswordAuto = "1";
        parameterList.addParameter("passwordAuto", "int", tempPasswordAuto);
        string tempPasswordMandatory = "0";
        if (passwordMandatory)
            tempPasswordMandatory = "1";
        parameterList.addParameter("passwordMandatory", "int", tempPasswordMandatory);
        string tempPasswordOptional = "0";
        if (passwordOptional)
            tempPasswordOptional = "1";
        parameterList.addParameter("passwordOptional", "int", tempPasswordOptional);
        string tempHistoryAuto = "0";
        if (historyAuto)
            tempHistoryAuto = "1";
        parameterList.addParameter("historyAuto", "int", tempHistoryAuto);
        string tempHistoryOptional = "0";
        if (historyOptional)
            tempHistoryOptional = "1";
        parameterList.addParameter("historyOptional", "int", tempHistoryOptional);
        SqlDB dataReader = new SqlDB(SessionObj.Project.ProjectID);
        dataReader.execSQLwithParameter("UPDATE teamspace_container SET displayName=@displayName, maxSizeGB=@maxSizeGb, maxFileSizeMB=@maxFileSizeMB, passwordAuto=@passwordAuto, passwordMandatory=@passwordMandatory, passwordOptional=@passwordOptional, historyAuto=@historyAuto, historyOptional=@historyOptional WHERE containerID=@containerID", parameterList);
    }
    public static void add(int aConatinerID, string aDisplayName, int aMaxSizeGB, int aMaxFileSizeMB, bool aPasswortAuto, bool aPasswortMandatory, bool aPasswordOptional, bool aHistoryAuto, bool aHistoryOptional)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("containerID", "int", aConatinerID.ToString());
        parameterList.addParameter("displayName", "string", aDisplayName);
        parameterList.addParameter("maxSizeGB", "int", aMaxSizeGB.ToString());
        parameterList.addParameter("maxFileSizeMB", "int", aMaxFileSizeMB.ToString());
        string tempPasswordAuto = "0";
        if (aPasswortAuto)
            tempPasswordAuto = "1";
        parameterList.addParameter("passwordAuto", "int", tempPasswordAuto);
        string tempPasswordMandatory = "0";
        if (aPasswortMandatory)
            tempPasswordMandatory = "1";
        parameterList.addParameter("passwordMandatory", "int", tempPasswordMandatory);
        string tempPasswordOptional = "0";
        if (aPasswordOptional)
            tempPasswordOptional = "1";
        parameterList.addParameter("passwordOptional", "int", tempPasswordOptional);
        string tempHistoryAuto = "0";
        if (aHistoryAuto)
            tempHistoryAuto = "1";
        parameterList.addParameter("historyAuto", "int", tempHistoryAuto);
        string tempHistoryOptional = "0";
        if (aHistoryOptional)
            tempHistoryOptional = "1";
        parameterList.addParameter("historyOptional", "int", tempHistoryOptional);
        SqlDB dataReader = new SqlDB(SessionObj.Project.ProjectID);
        //try
        {
            dataReader.execSQLwithParameter("INSERT INTO teamspace_container (displayName, maxSizeGB, maxFileSizeMB, passwordAuto, passwordMandatory, passwordOptional, historyAuto, historyOptional) values (@displayName, @maxSizeGB, @maxFileSizeMB, @passwordAuto, @passwordMandatory, @passwordOptional, @historyAuto, @historyOptional)", parameterList);
        }
        //catch
        {
        }
    }
    public static ListItemCollection getContainerList()
    {
        ListItemCollection result = new ListItemCollection();

        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        SqlDB dataReader;
        dataReader = new SqlDB("SELECT containerID, displayName FROM teamspace_container", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            ListItem tempItem = new ListItem(dataReader.getString(1), dataReader.getInt32(0).ToString());
            result.Add(tempItem);
        }
        dataReader.close();

        return result;
    }
}